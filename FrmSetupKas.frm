VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{30A75567-DC01-443C-B1C3-82A76E79E5C9}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmSetupKas 
   Caption         =   "Master Kas"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSetupKas.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6840
   ScaleWidth      =   9000
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
      Height          =   6090
      Left            =   120
      ScaleHeight     =   6030
      ScaleWidth      =   8790
      TabIndex        =   7
      Top             =   0
      Width           =   8850
      Begin VB.PictureBox Picture2 
         Height          =   5010
         Left            =   150
         ScaleHeight     =   4950
         ScaleWidth      =   8430
         TabIndex        =   8
         Top             =   870
         Width           =   8490
         Begin MSDataGridLib.DataGrid DgMaster 
            Bindings        =   "FrmSetupKas.frx":08CA
            Height          =   2625
            Left            =   105
            TabIndex        =   5
            Tag             =   "KAS"
            Top             =   2190
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   4630
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
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
               DataField       =   "BankID"
               Caption         =   "Kode Kas"
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
               DataField       =   "TypeBank"
               Caption         =   "Jenis Kas"
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
               DataField       =   "NamaBank"
               Caption         =   "Nama Kas"
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
                  ColumnWidth     =   2055.118
               EndProperty
               BeginProperty Column02 
                  DividerStyle    =   6
                  ColumnWidth     =   3555.213
               EndProperty
            EndProperty
         End
         Begin VB.TextBox txtKas 
            Appearance      =   0  'Flat
            DataField       =   "Amount"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0;(#,##0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            DataSource      =   "DataMaster"
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   2055
            MaxLength       =   15
            TabIndex        =   4
            Tag             =   "KAS"
            Top             =   1575
            Width           =   3555
         End
         Begin VB.TextBox txtKas 
            Appearance      =   0  'Flat
            DataField       =   "NamaBank"
            DataSource      =   "DataMaster"
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   2055
            MaxLength       =   50
            TabIndex        =   3
            Tag             =   "KAS"
            Top             =   1245
            Width           =   3555
         End
         Begin VB.TextBox txtKas 
            Appearance      =   0  'Flat
            DataField       =   "TypeBank"
            DataSource      =   "DataMaster"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   2055
            MaxLength       =   50
            TabIndex        =   2
            Tag             =   "KAS"
            Top             =   915
            Width           =   3555
         End
         Begin VB.TextBox txtKas 
            Appearance      =   0  'Flat
            DataField       =   "BankID"
            DataSource      =   "DataMaster"
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   2055
            MaxLength       =   10
            TabIndex        =   0
            Tag             =   "KAS"
            Top             =   270
            Width           =   1500
         End
         Begin MSDataListLib.DataCombo cboKas 
            DataField       =   "NoAccount"
            DataSource      =   "DataMaster"
            Height          =   330
            Left            =   2055
            TabIndex        =   1
            Tag             =   "KAS"
            Top             =   570
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            Appearance      =   0
            Style           =   2
            BackColor       =   -2147483643
            ListField       =   "AccountName"
            BoundColumn     =   "NoAccount"
            Text            =   ""
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   900
            X2              =   2160
            Y1              =   1875
            Y2              =   1875
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   900
            X2              =   2160
            Y1              =   1545
            Y2              =   1545
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   900
            X2              =   2160
            Y1              =   1215
            Y2              =   1215
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   900
            X2              =   2160
            Y1              =   885
            Y2              =   885
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   900
            X2              =   2160
            Y1              =   540
            Y2              =   540
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Awal"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   2
            Left            =   930
            TabIndex        =   13
            Top             =   1605
            Width           =   915
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Kas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   1
            Left            =   930
            TabIndex        =   12
            Top             =   1275
            Width           =   825
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
            BackStyle       =   0  'Transparent
            Caption         =   "Jenis Kas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   0
            Left            =   930
            TabIndex        =   11
            Top             =   945
            Width           =   780
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
            BackStyle       =   0  'Transparent
            Caption         =   "No. Kas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   29
            Left            =   930
            TabIndex        =   10
            Top             =   300
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
            BackStyle       =   0  'Transparent
            Caption         =   "No. Akun."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   28
            Left            =   930
            TabIndex        =   9
            Top             =   630
            Width           =   780
         End
      End
   End
   Begin SemeruDC.SemeruOleDC Mydde 
      Align           =   2  'Align Bottom
      Height          =   690
      Left            =   0
      TabIndex        =   6
      Top             =   6150
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   1217
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmSetupKas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RcAcc As New DBQuick

Private Sub cboKas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
OpenAccount
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
OpenAccount
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmSetupKas
    .BindFormTAG = "KAS"
    Set .ActiveConnection = Cnn
    .PrepareQuery = "SELECT BankID, NoAccount, TypeBank, NamaBank, Amount FROM         [Temp Bank] where Typetrans ='KAS' ORDER BY BankID"
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MyDDE.CheckRecordPendinged = True Then
   ScanKey vbKeyF5, 0, MyDDE
   If MyDDE.IsSucces = True Then
      Cancel = False
      MyDDE.ClearRecordset
      RcAcc.CloseDB
   Else
      Cancel = True
   End If
Else
   MyDDE.ClearRecordset
   RcAcc.CloseDB
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
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbAddNew:
            'mVarDataDc = True
            txtKas(0).SetFocus
            MyDDE.GetFieldByName("Amount") = 0
       Case tmbEdit:
            txtKas(0).Enabled = False
            txtKas(3).Enabled = False
            'mVarDataDc = True
            txtKas(1).SetFocus
       Case tmbPrint:
            CallRPTReport "Laporan Kas.rpt"
       Case Else: 'mVarDataDc = False
End Select
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim mDel As New clsDelete
Select Case AdReasonActiveDb
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               If mDel.CekDelete(txtKas(0), reDelMasterCurency) = False Then
                  MyDDE.IsChildMemberReady = True
                  
               Else
                  MyDDE.CancelTrans = True
                  MessageBox "Record (" & txtKas(0) & ") Sedang Dipakai Transaksi Lain." & vbCrLf & "Record Tidak Bisa DiHapus.", "Peringatan", msgOkOnly
                  MyDDE.IsChildMemberReady = False
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
Set mDel = Nothing
End Sub

Private Sub txtKas_GotFocus(Index As Integer)
Block txtKas(Index)
End Sub

Private Sub txtKas_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub PrepareQuery()
On Error Resume Next
With MyDDE
    .PrepareAppend = " INSERT INTO [Temp Bank]" & _
                     " (BankID, NoAccount, TypeBank, NamaBank, Amount,TypeTrans) " & _
                     " VALUES     (N'" & ValidString(txtKas(0)) & "', N'" & cboKas.BoundText & "', N'" & ValidString(txtKas(1)) & "', N'" & ValidString(txtKas(2)) & "', " & CDbl(txtKas(3)) & ",'KAS')"
    .PrepareUpdate = " UPDATE [Temp Bank] Set [TypeBank] = N'" & ValidString(txtKas(1)) & "',[NamaBank] = N'" & ValidString(txtKas(2)) & "',Amount=" & CDbl(txtKas(3)) & " WHERE     (BankID = N'" & ValidString(txtKas(0)) & "')"
    .PrepareDelete = " DELETE FROM [Temp Bank] WHERE   (BankID = N'" & ValidString(txtKas(0)) & "') "
End With
Err.Clear
End Sub

Private Sub OpenAccount()
RcAcc.DBOpen "SELECT     NoAccount, AccountName FROM         GlAccount WHERE     (Type = N'Cash') AND ([Group] = N'List Account') ORDER BY AccountName", Cnn, lckLockReadOnly
Set cboKas.RowSource = RcAcc.DBRecordset
End Sub

Private Sub txtKas_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 3 Then ValidNum KeyAscii
End Sub


