VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmResource 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resource"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9180
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmResource.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   9180
   Tag             =   "Resources"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   4515
      Left            =   0
      ScaleHeight     =   4515
      ScaleWidth      =   9180
      TabIndex        =   5
      Top             =   0
      Width           =   9180
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Keterangan"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   1
         Left            =   1515
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "Partner"
         Top             =   450
         Width           =   3045
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Resources ID"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   0
         Left            =   1515
         MaxLength       =   15
         TabIndex        =   1
         Tag             =   "Partner"
         Top             =   105
         Width           =   1935
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Catatan"
         DataSource      =   "MyDDE"
         Height          =   660
         Index           =   2
         Left            =   1515
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Tag             =   "Partner"
         Top             =   1140
         Width           =   5430
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         DataField       =   "TypeID"
         DataSource      =   "MyDDE"
         Height          =   330
         Left            =   1515
         TabIndex        =   3
         Tag             =   "Partner"
         Top             =   795
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Description"
         BoundColumn     =   "TypeID"
         Text            =   "DataCombo1"
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2490
         Index           =   0
         Left            =   195
         TabIndex        =   6
         Tag             =   "Partner"
         Top             =   1890
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   4392
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
            DataField       =   "Resources ID"
            Caption         =   "Resources ID"
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
            DataField       =   "TypeID"
            Caption         =   "Type ID"
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
         Index           =   2
         X1              =   1600
         X2              =   195
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   870
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   510
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resources ID"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   165
         Width           =   1065
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   3300
         X2              =   195
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   1
         X1              =   4395
         X2              =   195
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   3
         X1              =   4395
         X2              =   195
         Y1              =   1785
         Y2              =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Catatan"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1545
         Width           =   630
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   4530
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RcType As New DBQuick

Private Sub DataCombo1_Click(Area As Integer)
If txtBox(1).Enabled = True Then
   MyDDE.GetFieldByName("TypeID") = DataCombo1.BoundText
End If
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
OpenType
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmResource
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT  [Resources Table].ResourcesID AS [Resources ID], [Resources Table].Description AS Keterangan, [Resources Table].TypeID,[Resources Type].Description AS [Type Desc], [Resources Table].Note AS Catatan FROM         [Resources Table] INNER JOIN                       [Resources Type] ON [Resources Table].TypeID = [Resources Type].TypeID"
End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MyDDE.CheckRecordPendinged = True Then
   ScanKey vbKeyF5, 0, MyDDE
   If MyDDE.IsSucces = True Then
      Cancel = False
      MyDDE.ClearRecordset
      Set FrmResource = Nothing
   Else
      Cancel = True
   End If
Else
   MyDDE.ClearRecordset
   Set FrmResource = Nothing
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmResource = Nothing
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbAddNew:
            'mVarDataDc = True
            txtBox(0).SetFocus
            MyDDE.GetFieldByName("Catatan") = "-"
       Case tmbEdit:
            txtBox(0).Enabled = False
            'mVarDataDc = True
            txtBox(1).SetFocus
       Case tmbPrint:
            CallRPTReport "Resources Man Type.rpt"
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
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub PrepareQuery()
With MyDDE
    .PrepareAppend = " INSERT INTO [Resources Table] ([ResourcesID], [Description],[TypeID],Note) " & _
                     " VALUES (N'" & (txtBox(0)) & "', N'" & txtBox(1) & "',N'" & DataCombo1.BoundText & "',N'" & txtBox(2) & "')"
                     
    .PrepareUpdate = " UPDATE [Resources Table] Set [Description] = N'" & txtBox(1) & "',[TypeID]=N'" & DataCombo1.BoundText & "',Note=N'" & txtBox(2) & "' WHERE     ([ResourcesID] = N'" & txtBox(0) & "')"
    
    .PrepareDelete = " DELETE FROM [Resources Table] WHERE   ([ResourcesID] = N'" & txtBox(0) & "') "
End With
End Sub

Private Sub GridLayout()
DataGrid1(0).Height = 2375
DataGrid1(0).width = 8400
DataGrid1(0).Columns(0).width = 2250.142
DataGrid1(0).Columns(1).width = 3314.835
DataGrid1(0).Columns(2).width = 2280.189
End Sub

Private Sub OpenType()
RcType.DBOpen "SELECT     * FROM         [Resources Type] ORDER BY Description", CNN, lckLockReadOnly
Set DataCombo1.RowSource = RcType.DBRecordset
End Sub







