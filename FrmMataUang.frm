VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5BD1BD0-C880-4C3C-8176-E61FC2E2B3F5}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmMataUang 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8865
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMataUang.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   Tag             =   "Exchange Rate Maintenance"
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
      Height          =   5025
      Left            =   0
      ScaleHeight     =   4995
      ScaleWidth      =   8730
      TabIndex        =   7
      Top             =   0
      Width           =   8760
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Height          =   4275
         Left            =   135
         ScaleHeight     =   4245
         ScaleWidth      =   8445
         TabIndex        =   8
         Top             =   540
         Width           =   8475
         Begin VB.TextBox txtBox 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "Rate"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0;(#,##0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   5
            Tag             =   "Partner"
            Top             =   855
            Width           =   1935
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   2865
            Index           =   0
            Left            =   180
            TabIndex        =   6
            Tag             =   "Partner"
            Top             =   1230
            Width           =   8130
            _ExtentX        =   14340
            _ExtentY        =   5054
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
               Weight          =   700
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
               DataField       =   "CurrID"
               Caption         =   "Currency"
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
               DataField       =   "Currency Name"
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
            BeginProperty Column02 
               DataField       =   "Rate"
               Caption         =   "Rate"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0;(#,##0)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               BeginProperty Column00 
                  DividerStyle    =   6
                  ColumnWidth     =   2280,189
               EndProperty
               BeginProperty Column01 
                  DividerStyle    =   6
                  ColumnWidth     =   3885,166
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   1409,953
               EndProperty
            EndProperty
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "CurrID"
            Height          =   270
            Index           =   0
            Left            =   1410
            MaxLength       =   5
            TabIndex        =   1
            Tag             =   "Partner"
            Top             =   255
            Width           =   1935
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "Currency Name"
            Height          =   270
            Index           =   1
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   3
            Tag             =   "Partner"
            Top             =   555
            Width           =   3165
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rate (Rp)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   2
            Left            =   225
            TabIndex        =   4
            Top             =   870
            Width           =   705
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000001&
            Index           =   2
            X1              =   3330
            X2              =   195
            Y1              =   1110
            Y2              =   1110
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000001&
            Index           =   1
            X1              =   4410
            X2              =   210
            Y1              =   810
            Y2              =   810
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   3315
            X2              =   210
            Y1              =   510
            Y2              =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Currency"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   0
            Top             =   255
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   2
            Top             =   570
            Width           =   795
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   9
      Top             =   5130
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmMataUang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmMataUang
    .BindFormTAG = "Partner"
    Set .ActiveConnection = Cnn
    .PrepareQuery = "Select * from [Currency Table]"
End With
HiasForm Picture1, Me
CenterForm Picture2, Me
GridLayout
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Me.hwnd
End Sub

Private Sub Form_Resize()
'HiasForm Picture1, Me
'CenterForm Picture2, Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MyDDE.CheckRecordPendinged = True Then
   ScanKey vbKeyF5, 0, MyDDE
   If MyDDE.IsSucces = True Then
      Cancel = False
      MyDDE.ClearRecordset
      Set FrmMataUang = Nothing
   Else
      Cancel = True
   End If
Else
   MyDDE.ClearRecordset
   Set FrmMataUang = Nothing
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmMataUang = Nothing
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
            CallRPTReport "Tabel Mata Uang.rpt"
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
MoveForm Picture1.Parent.hwnd
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub PrepareQuery()
With MyDDE
    .PrepareAppend = " INSERT INTO [Currency Table] (CurrID, [Currency Name],[Rate]) " & _
                     " VALUES (N'" & ValidString(txtBox(0)) & "', N'" & ValidString(txtBox(1)) & "'," & CDbl(txtBox(2)) & ")"
                     
    .PrepareUpdate = " UPDATE [Currency Table] Set [Currency Name] = N'" & txtBox(1) & "' , Rate =" & CDbl(txtBox(2)) & " WHERE     (CurrID = N'" & ValidString(txtBox(0)) & "')"
    
    .PrepareDelete = " DELETE FROM [Currency Table] WHERE   (CurrID = N'" & ValidString(txtBox(0)) & "') "
End With
End Sub

Private Sub GridLayout()
DataGrid1(0).Columns(0).Width = 2280.189
DataGrid1(0).Columns(1).Width = 3885.166
DataGrid1(0).Columns(2).Width = 1409.953
End Sub

Private Sub txtBox_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 2 Then ValidNum KeyAscii
End Sub
