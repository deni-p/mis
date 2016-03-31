VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{B1E614FF-F86D-4F68-A86F-2584A0570C66}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmResourcesType 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resource"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmResourcesType.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   9150
   Tag             =   "Resources Type"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   3960
      Left            =   0
      ScaleHeight     =   3960
      ScaleWidth      =   9150
      TabIndex        =   1
      Top             =   0
      Width           =   9150
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Resouces ID"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   0
         Left            =   1380
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Partner"
         Top             =   210
         Width           =   1935
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Keterangan"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   1
         Left            =   1380
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "Partner"
         Top             =   570
         Width           =   3045
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Time Used"
         DataField       =   "Waktu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   180
         TabIndex        =   2
         Tag             =   "Partner"
         Top             =   930
         Width           =   1380
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2550
         Index           =   0
         Left            =   195
         TabIndex        =   5
         Tag             =   "Partner"
         Top             =   1245
         Width           =   8790
         _ExtentX        =   15505
         _ExtentY        =   4498
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
            DataField       =   "Resouces ID"
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
            DataField       =   "Waktu"
            Caption         =   "Time Used"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "YES"
               FalseValue      =   "NO"
               NullValue       =   "NO"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
               DividerStyle    =   6
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   6
               ColumnWidth     =   3885.166
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   1934.929
            EndProperty
         EndProperty
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   1
         X1              =   4380
         X2              =   180
         Y1              =   870
         Y2              =   870
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   3285
         X2              =   180
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resouces ID"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   0
         Left            =   210
         TabIndex        =   7
         Top             =   255
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   1
         Left            =   210
         TabIndex        =   6
         Top             =   615
         Width           =   900
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   3975
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmResourcesType"
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
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmResourcesType
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT     TypeID AS [Resouces ID], Description AS Keterangan, TimeUsed AS Waktu FROM         [Resources Type] ORDER BY TypeID"
End With

End Sub

Private Sub Form_Resize()

'GridLayout
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MyDDE.CheckRecordPendinged = True Then
   ScanKey vbKeyF5, 0, MyDDE
   If MyDDE.IsSucces = True Then
      Cancel = False
      MyDDE.ClearRecordset
      Set FrmResourcesType = Nothing
   Else
      Cancel = True
   End If
Else
   MyDDE.ClearRecordset
   Set FrmResourcesType = Nothing
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmResourcesType = Nothing
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbAddNew:
            'mVarDataDc = True
            txtBox(0).SetFocus
       Case tmbEdit:
            txtBox(0).Enabled = False
            'mVarDataDc = True
            txtBox(1).SetFocus
       Case tmbPrint:
            CallRPTReport "Resources Type Table.rpt"
       Case Else: 'mVarDataDc = False
End Select
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim mDel As New clsDelete
Select Case AdReasonActiveDb
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               If mDel.CekDelete(txtBox(0), reDelMasterResources) = False Then
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
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub PrepareQuery()
With MyDDE
    .PrepareAppend = " INSERT INTO [Resources Type] ([TypeID], [Description],[TimeUsed]) " & _
                     " VALUES (N'" & (txtBox(0)) & "', N'" & txtBox(1) & "'," & Check1.value & ")"
                     
    .PrepareUpdate = " UPDATE [Resources Type] Set [Description] = N'" & txtBox(1) & "',TimeUsed=" & Check1.value & " WHERE     ([TypeID] = N'" & txtBox(0) & "')"
    
    .PrepareDelete = " DELETE FROM [Resources Type] WHERE   ([TypeID] = N'" & txtBox(0) & "') "
End With
End Sub

Private Sub GridLayout()
DataGrid1(0).Height = 2225
DataGrid1(0).Width = 8475
DataGrid1(0).Columns(0).Width = 2085.166
DataGrid1(0).Columns(1).Width = 3885.166
DataGrid1(0).Columns(2).Width = 1934.929
End Sub







