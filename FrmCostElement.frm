VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmCostElement 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Konfigurasi Biaya"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9045
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCostElement.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   9045
   Tag             =   "Cost Methode"
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   6
      Top             =   4020
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   4065
      Left            =   0
      ScaleHeight     =   4065
      ScaleWidth      =   9045
      TabIndex        =   7
      Top             =   0
      Width           =   9045
      Begin VB.CommandButton CmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   7185
         Picture         =   "FrmCostElement.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "CURR"
         Top             =   1223
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Tipe Cost"
         DataSource      =   "Adodc1"
         Height          =   330
         Index           =   0
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Partner"
         Top             =   135
         Width           =   3015
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Keterangan"
         DataSource      =   "Adodc1"
         Height          =   330
         Index           =   1
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "Partner"
         Top             =   495
         Width           =   3015
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2310
         Index           =   0
         Left            =   195
         TabIndex        =   5
         Tag             =   "Partner"
         Top             =   1620
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   4075
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16577005
         ForeColor       =   7159830
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   6
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Tipe Cost"
            Caption         =   "Cost Methode"
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
            DataField       =   "Keterangan"
            Caption         =   "Description"
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
               DividerStyle    =   6
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   6
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Description"
         DataField       =   "perkiraan"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   5
         Left            =   1680
         TabIndex        =   3
         Tag             =   "Partner"
         Top             =   1215
         Visible         =   0   'False
         Width           =   5505
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   1
         X1              =   2000
         X2              =   240
         Y1              =   810
         Y2              =   810
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2000
         X2              =   225
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Biaya"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   195
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   555
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1275
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   2
         Visible         =   0   'False
         X1              =   2000
         X2              =   240
         Y1              =   1530
         Y2              =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Rekening"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   915
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Kode"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   4
         Left            =   1680
         TabIndex        =   2
         Tag             =   "Partner"
         Top             =   855
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   3
         Visible         =   0   'False
         X1              =   2000
         X2              =   240
         Y1              =   1170
         Y2              =   1170
      End
   End
End
Attribute VB_Name = "FrmCostElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RcKelompok As New Recordset
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Dim strSQL As String

Private Sub cmdLink_Click(Index As Integer)
'   RcKelompok.Open "SELECT     GLAccount.NoAccount AS Kode, GLAccount.AccountName AS Perkiraan FROM         GLAccount INNER JOIN AccType ON GLAccount.Type = AccType.Tipe WHERE     (AccType.ID = N'0' OR AccType.ID = N'15' OR AccType.ID = N'52' OR AccType.ID = N'53') AND (GLAccount.[Group] = N'Detail List Account')", CNN, adOpenForwardOnly, adLockReadOnly
strSQL = "SELECT     GLAccount.NoAccount AS Kode, GLAccount.AccountName AS Perkiraan " & _
    " FROM GLAccount INNER JOIN AccType ON GLAccount.Type = AccType.Tipe " & _
    " WHERE (AccType.ID = N'72' OR AccType.ID = N'15' OR AccType.ID = N'52' OR AccType.ID = N'53') " & _
    " AND (GLAccount.[Group] = N'Detail List Account')"

RcKelompok.Open strSQL, CNN, adOpenForwardOnly, adLockReadOnly
   
Set mCall.FormData = RcKelompok
mCall.CaptionLink = "Account"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
GridLayout
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmCostElement
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
    '.PrepareQuery = "SELECT [Cost Element].[Cost Element Type] AS [Tipe Cost], [Cost Element].Description AS Keterangan, [Cost Element].NoAccount AS Kode,  GLAccount.AccountName AS Perkiraan FROM         [Cost Element] INNER JOIN GLAccount ON [Cost Element].NoAccount = GLAccount.NoAccount ORDER BY [Cost Element].[Cost Element Type]"
    .PrepareQuery = "SELECT [Cost Element Type] AS [Tipe Cost], Description AS Keterangan From [Cost Element] ORDER BY [Tipe Cost]"
'Debug.Print .PrepareQuery
End With
Set mCall = New frmCaller
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MyDDE.CheckRecordPendinged = True Then
   ScanKey vbKeyF5, 0, MyDDE
   If MyDDE.IsSucces = True Then
      Cancel = False
      MyDDE.ClearRecordset
      Set FrmCostElement = Nothing
      RcKelompok.Close
      Set RcKelompok = Nothing
   Else
      Cancel = True
   End If
Else
   MyDDE.ClearRecordset
   Set FrmCostElement = Nothing
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmCostElement = Nothing
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
   MyDDE.GetFieldByName("perkiraan") = mCall.GetFieldByName("perkiraan")
   MyDDE.GetFieldByName("kode") = mCall.GetFieldByName("kode")
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 1
cmdLink(0).Enabled = False
Select Case AdReasonActiveDb
       Case tmbAddNew:
            'mVarDataDc = True
            txtBox(0).SetFocus
            cmdLink(0).Enabled = True
       Case tmbEdit:
            txtBox(0).Enabled = False
            'mVarDataDc = True
            txtBox(1).SetFocus
            cmdLink(0).Enabled = True
       Case tmbPrint:
            CallRPTReport "Cost Element Table.rpt"
       Case Else: 'mVarDataDc = False
End Select
Exit Sub
1:
MessageBox Err.Description, "frmcostelement:mydde_afterprepareactivedb & Err.Number, msgOkOnly, msgExclamation"
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 2
Dim mDel As New clsDelete
Select Case AdReasonActiveDb
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               If mDel.CekDelete(txtBox(0), reDelMasterCurency) = False Then
                  MyDDE.IsChildMemberReady = True
                  PrepareQuery
               Else
                  MyDDE.CancelTrans = True
                  MessageBox "Record (" & txtBox(0) & ") Sedang Dipakai Transaksi Lain." & vbCrLf & "Record Tidak Bisa DiHapus.", "Peringatan", msgOkOnly, msgCrtical
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
Exit Sub
2:
MessageBox Err.Description, "frmcostelement:mydde_prepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub PrepareQuery()
On Error GoTo xErr
With MyDDE
    .PrepareAppend = " INSERT INTO [Cost Element] ([Cost Element Type], [Description]) " & _
                     " VALUES (N'" & (txtBox(0)) & "', N'" & txtBox(1) & "')"
                     
    .PrepareUpdate = " UPDATE [Cost Element] Set [Description] = N'" & txtBox(1) & "' WHERE     ([Cost Element Type] = N'" & txtBox(0) & "')"
    
    .PrepareDelete = " DELETE FROM [Cost Element] WHERE   ([Cost Element Type] = N'" & txtBox(0) & "') "
End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub GridLayout()
DataGrid1(0).Height = 2225
DataGrid1(0).width = 8280
DataGrid1(0).Columns(0).width = 3089.764
DataGrid1(0).Columns(1).width = 4619.906
End Sub





