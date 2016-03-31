VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmUOM 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Satuan Barang"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8850
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmUOM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   8850
   Tag             =   "Unit of Measurement"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4605
      Left            =   0
      ScaleHeight     =   4605
      ScaleWidth      =   8850
      TabIndex        =   3
      Top             =   0
      Width           =   8850
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "uom"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   0
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   1
         Tag             =   "Partner"
         Top             =   195
         Width           =   1935
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "description"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   1
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "Partner"
         Top             =   540
         Width           =   3045
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmUOM.frx":6852
         Height          =   2985
         Index           =   0
         Left            =   195
         TabIndex        =   4
         Tag             =   "Partner"
         Top             =   1290
         Width           =   8400
         _ExtentX        =   14817
         _ExtentY        =   5265
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16777215
         HeadLines       =   2
         RowHeight       =   15
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "UOM"
            Caption         =   "UOM"
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
            DataField       =   "Description"
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
            MarqueeStyle    =   3
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
         Caption         =   "Keterangan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   1
         Left            =   195
         TabIndex        =   6
         Top             =   585
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   0
         Left            =   195
         TabIndex        =   5
         Top             =   240
         Width           =   330
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1725
         X2              =   150
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   1725
         X2              =   150
         Y1              =   840
         Y2              =   840
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   4605
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmUOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()

MyDDE.SetPermissions = aksess.MayDo("Unit Of Measurement")
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
GridLayout
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmUOM
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT * FROM [UOM Table] ORDER BY uom"
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MyDDE.CheckRecordPendinged = True Then
   ScanKey vbKeyF5, 0, MyDDE
   If MyDDE.IsSucces = True Then
      Cancel = False
      MyDDE.ClearRecordset
      Set FrmCategories = Nothing
   Else
      Cancel = True
   End If
Else
   MyDDE.ClearRecordset
   Set FrmCategories = Nothing
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmCategories = Nothing
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
            CallRPTReport "Manufacture Inv Categories.rpt"
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
With MyDDE
    .PrepareAppend = " INSERT INTO [UOM Table] ([UOM], [Description]) " & _
                     " VALUES (N'" & (txtBox(0)) & "', N'" & txtBox(1) & "')"
                     
    .PrepareUpdate = " UPDATE [UOM Table] Set [Description] = N'" & txtBox(1) & "' WHERE ([UOM] = N'" & txtBox(0) & "')"
    
    .PrepareDelete = " DELETE FROM [UOM Table] WHERE   ([UOM] = N'" & txtBox(0) & "') "
End With
End Sub

Private Sub GridLayout()
DataGrid1(0).Height = 2850
DataGrid1(0).width = 8025
DataGrid1(0).Columns(0).width = 1934.929
DataGrid1(0).Columns(1).width = 5534.929
End Sub









