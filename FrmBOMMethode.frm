VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmBOMMethode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BOM Method"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9075
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBOMMethode.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   Tag             =   "BOM Methode"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   3570
      Left            =   0
      ScaleHeight     =   3570
      ScaleWidth      =   9075
      TabIndex        =   4
      Top             =   0
      Width           =   9075
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Keterangan"
         DataSource      =   "Adodc1"
         Height          =   330
         Index           =   1
         Left            =   1410
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "Partner"
         Top             =   570
         Width           =   3045
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Metode ID"
         DataSource      =   "Adodc1"
         Height          =   330
         Index           =   0
         Left            =   1410
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Partner"
         Top             =   195
         Width           =   1935
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2310
         Index           =   0
         Left            =   195
         TabIndex        =   3
         Tag             =   "Partner"
         Top             =   990
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   4075
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16577005
         ForeColor       =   7159830
         HeadLines       =   2
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
            DataField       =   "Metode ID"
            Caption         =   "Method ID"
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
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   645
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Method ID"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   255
         Width           =   750
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   3315
         X2              =   210
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   1
         X1              =   4410
         X2              =   210
         Y1              =   885
         Y2              =   885
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   3570
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmBOMMethode"
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
    Set .BindForm = FrmBOMMethode
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT     MethodeID AS [Metode ID], Description AS Keterangan FROM         [BOM Methode] ORDER BY MethodeID"
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo 1
If MyDDE.CheckRecordPendinged = True Then
   ScanKey vbKeyF5, 0, MyDDE
   If MyDDE.IsSucces = True Then
      Cancel = False
      MyDDE.ClearRecordset
      Set FrmBOMMethode = Nothing
   Else
      Cancel = True
   End If
Else
   MyDDE.ClearRecordset
   Set FrmBOMMethode = Nothing
End If
Exit Sub
1:
MessageBox Err.Description, "frmbommethode:form_queryunload_" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmBOMMethode = Nothing
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
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
            CallRPTReport "Bom Methode Table.rpt"
       Case Else: 'mVarDataDc = False
End Select
Exit Sub
1:
MessageBox Err.Description, "frmbommethode:mydde_afterprepareactivedb" & Err.Number, msgOkOnly, msgExclamation
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
MessageBox Err.Description, "frmbommethode:mydde_prepareactivedb" & Err.Number, msgOkOnly, msgExclamation
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
    .PrepareAppend = " INSERT INTO [BOM Methode] ([MethodeID], [Description]) " & _
                     " VALUES (N'" & (txtBox(0)) & "', N'" & txtBox(1) & "')"
                     
    .PrepareUpdate = " UPDATE [BOM Methode] Set [Description] = N'" & txtBox(1) & "' WHERE     ([MethodeID] = N'" & txtBox(0) & "')"
    
    .PrepareDelete = " DELETE FROM [BOM Methode] WHERE   ([MethodeID] = N'" & txtBox(0) & "') "
End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub GridLayout()
DataGrid1(0).Height = 2225
DataGrid1(0).width = 8475
DataGrid1(0).Columns(0).width = 1904.882
DataGrid1(0).Columns(1).width = 6015.118
End Sub





