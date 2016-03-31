VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{30A75567-DC01-443C-B1C3-82A76E79E5C9}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmListingFixAssets 
   Caption         =   "Daftar Group Aktiva Tetap"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListingFixAssets.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   9255
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   4815
      Left            =   0
      ScaleHeight     =   4755
      ScaleWidth      =   9990
      TabIndex        =   4
      Top             =   0
      Width           =   10050
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   3450
         Left            =   375
         ScaleHeight     =   3420
         ScaleWidth      =   7290
         TabIndex        =   5
         Top             =   660
         Width           =   7320
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   2250
            Index           =   0
            Left            =   210
            TabIndex        =   2
            Tag             =   "Partner"
            Top             =   1005
            Width           =   6780
            _ExtentX        =   11959
            _ExtentY        =   3969
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   16777215
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
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   "Id Group"
               Caption         =   "Id Group"
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
               DataField       =   "Aktiva Group"
               Caption         =   "Aktiva Group"
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
                  ColumnWidth     =   1814.74
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   4364.788
               EndProperty
            EndProperty
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "Aktiva Group"
            Height          =   315
            Index           =   1
            Left            =   1965
            MaxLength       =   50
            TabIndex        =   1
            Tag             =   "Partner"
            Top             =   570
            Width           =   3045
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "Id Group"
            Height          =   315
            Index           =   0
            Left            =   1965
            MaxLength       =   5
            TabIndex        =   0
            Tag             =   "Partner"
            Top             =   210
            Width           =   1935
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   630
            X2              =   2010
            Y1              =   870
            Y2              =   870
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   645
            X2              =   2025
            Y1              =   510
            Y2              =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aktiva Group"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   1
            Left            =   660
            TabIndex        =   7
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Id Group"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   0
            Left            =   660
            TabIndex        =   6
            Top             =   255
            Width           =   810
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   690
      Left            =   0
      TabIndex        =   3
      Top             =   5460
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   1217
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmListingFixAssets"
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
    Set .BindForm = frmListingFixAssets
    .BindFormTAG = "Partner"
    Set .ActiveConnection = Cnn
    .PrepareQuery = "SELECT     [Id Group], [Aktiva Group] FROM         [Aktiva Class] ORDER BY [Id Group]"
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MyDDE.CheckRecordPendinged = True Then
   ScanKey vbKeyF5, 0, MyDDE
   If MyDDE.IsSucces = True Then
      Cancel = False
      MyDDE.ClearRecordset
   Else
      Cancel = True
   End If
Else
   MyDDE.ClearRecordset
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState <> vbMaximized Then
   Me.Height = MainMenu.ScaleHeight
   Me.Width = MainMenu.ScaleWidth
End If
HiasForm Picture1, Me
CenterForm Picture2
GridLayout
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
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

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub


Private Sub PrepareQuery()
With MyDDE
    .PrepareAppend = " INSERT INTO  [Aktiva Class] ( [Id Group], [Aktiva Group]) " & _
                     " VALUES (N'" & ValidString(txtBox(0)) & "', N'" & ValidString(txtBox(1)) & "')"

    .PrepareUpdate = " UPDATE  [Aktiva Class] Set [Aktiva Group] = N'" & ValidString(txtBox(1)) & "' WHERE     ( [Id Group] = N'" & ValidString(txtBox(0)) & "')"

    .PrepareDelete = " DELETE FROM  [Aktiva Class] WHERE   ( [Id Group] = N'" & ValidString(txtBox(0)) & "') "
End With
End Sub

Private Sub GridLayout()
DataGrid1(0).Height = 2250
DataGrid1(0).Width = 6780
End Sub


