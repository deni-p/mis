VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmDescriptor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descriptor"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9060
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmDescriptor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   9060
   Tag             =   "Outsourced Type"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   0
      ScaleHeight     =   3975
      ScaleWidth      =   9060
      TabIndex        =   5
      Top             =   0
      Width           =   9060
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Tipe ID"
         DataSource      =   "Adodc1"
         Height          =   330
         Index           =   0
         Left            =   1470
         MaxLength       =   15
         TabIndex        =   1
         Tag             =   "Partner"
         Top             =   210
         Width           =   1935
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Keterangan"
         DataSource      =   "Adodc1"
         Height          =   330
         Index           =   1
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "Partner"
         Top             =   555
         Width           =   3045
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "Klasifikasi"
         Height          =   330
         ItemData        =   "FrmDescriptor.frx":6852
         Left            =   1470
         List            =   "FrmDescriptor.frx":685F
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Partner"
         Top             =   900
         Width           =   3045
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2520
         Index           =   0
         Left            =   135
         TabIndex        =   4
         Tag             =   "Partner"
         Top             =   1320
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   4445
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16577005
         ForeColor       =   7159830
         HeadLines       =   2
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
            DataField       =   "Tipe ID"
            Caption         =   "Tipe ID"
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
         BeginProperty Column02 
            DataField       =   "Klasifikasi"
            Caption         =   "Klasifikasi"
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
            BeginProperty Column02 
            EndProperty
         EndProperty
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   1
         X1              =   4335
         X2              =   135
         Y1              =   870
         Y2              =   870
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   3240
         X2              =   135
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assembly Item"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   270
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   615
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Klasifikasi"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   2
         Left            =   180
         TabIndex        =   6
         Top             =   975
         Width           =   705
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   2
         X1              =   1795
         X2              =   135
         Y1              =   1215
         Y2              =   1215
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   4005
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmDescriptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
GridLayout
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmDescriptor
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT     TypeID AS [Tipe ID], Description AS Keterangan, Clasification AS Klasifikasi FROM         [Descriptor Type]"
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MyDDE.CheckRecordPendinged = True Then
   ScanKey vbKeyF5, 0, MyDDE
   If MyDDE.IsSucces = True Then
      Cancel = False
      MyDDE.ClearRecordset
      Set FrmDescriptor = Nothing
   Else
      Cancel = True
   End If
Else
   MyDDE.ClearRecordset
   Set FrmDescriptor = Nothing
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmDescriptor = Nothing
End Sub

Private Sub MYDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 1
Select Case AdReasonActiveDb
       Case tmbAddNew:
            'mVarDataDc = True
            txtBox(0).SetFocus
       Case tmbEdit:
            txtBox(0).Enabled = False
            'mVarDataDc = True
            txtBox(1).SetFocus
       Case tmbPrint:
            CallRPTReport "Descriptor Table.rpt"
       Case Else: 'mVarDataDc = False
End Select
Exit Sub
1:
MessageBox Err.Description, "frmdescriptor:mydde_afterprepareactivedb" & Err.Number, msgOkOnly, msgExclamation
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
MessageBox Err.Description, "frmdescriptor:mydde_prepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

Private Sub Picture2_Click()

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
    .PrepareAppend = " INSERT INTO [Descriptor Type] ([TypeID], [Description],Clasification) " & _
                     " VALUES (N'" & (txtBox(0)) & "', N'" & txtBox(1) & "',N'" & Combo1.Text & "')"
                     
    .PrepareUpdate = " UPDATE [Descriptor Type] Set [Description] = N'" & txtBox(1) & "',Clasification=N'" & Combo1.Text & "' WHERE     ([TypeID] = N'" & txtBox(0) & "')"
    
    .PrepareDelete = " DELETE FROM [Descriptor Type] WHERE   ([TypeID] = N'" & txtBox(0) & "') "
End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub GridLayout()
DataGrid1(0).Height = 2375 '2225
DataGrid1(0).width = 8370
DataGrid1(0).Columns(0).width = 1814.74
DataGrid1(0).Columns(1).width = 4364.788
DataGrid1(0).Columns(2).width = 1635.024
End Sub







