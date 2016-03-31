VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmCurrencySetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup Mata Uang"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8865
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCurrencySetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   Tag             =   "Currency Setup"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   5760
      Left            =   0
      ScaleHeight     =   5760
      ScaleWidth      =   8865
      TabIndex        =   9
      Top             =   0
      Width           =   8865
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Currency Name"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   1
         Left            =   1410
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "Partner"
         Top             =   495
         Width           =   2520
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "CurrID"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   0
         Left            =   1410
         MaxLength       =   5
         TabIndex        =   1
         Tag             =   "Partner"
         Top             =   135
         Width           =   1380
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Source"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   3
         Left            =   1410
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "Partner"
         Top             =   855
         Width           =   2520
      End
      Begin VB.TextBox txtBox 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "Rate"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   4
         Left            =   5700
         TabIndex        =   5
         Tag             =   "Partner"
         Top             =   517
         Width           =   1935
      End
      Begin VB.OptionButton OptCurrency 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pembagi"
         ForeColor       =   &H00404040&
         Height          =   225
         Index           =   1
         Left            =   6660
         TabIndex        =   7
         Top             =   915
         Width           =   915
      End
      Begin VB.OptionButton OptCurrency 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pengali"
         ForeColor       =   &H00404040&
         Height          =   225
         Index           =   0
         Left            =   5745
         TabIndex        =   6
         Top             =   915
         Value           =   -1  'True
         Width           =   840
      End
      Begin VB.CheckBox ChkDefault 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Mata Uang Utama"
         DataField       =   "Functional"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   5700
         TabIndex        =   4
         Top             =   240
         Width           =   1740
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmCurrencySetup.frx":6852
         Height          =   4275
         Index           =   0
         Left            =   75
         TabIndex        =   8
         Tag             =   "Partner"
         Top             =   1275
         Width           =   8730
         _ExtentX        =   15399
         _ExtentY        =   7541
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
            Weight          =   700
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "CurrID"
            Caption         =   "Currency"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Currency Name"
            Caption         =   "Keterangan"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Rate"
            Caption         =   "Rate"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Functional"
            Caption         =   "Default"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Yes"
               FalseValue      =   "No"
               NullValue       =   "No"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Activation"
            Caption         =   "Activation"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Source"
            Caption         =   "Rate Source"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "RVariance"
            Caption         =   "Rate Variance"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "CMethod"
            Caption         =   "Method"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   555
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mata Uang"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   195
         Width           =   780
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   1
         X1              =   1500
         X2              =   210
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   2
         X1              =   1500
         X2              =   195
         Y1              =   1155
         Y2              =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sumber"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   13
         Top             =   915
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fungsi"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   3
         Left            =   4155
         TabIndex        =   12
         Top             =   195
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kurs"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   4
         Left            =   4155
         TabIndex        =   11
         Top             =   555
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Metode Kalkulasi"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   5
         Left            =   4155
         TabIndex        =   10
         Top             =   915
         Width           =   1185
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   5760
         X2              =   4125
         Y1              =   810
         Y2              =   810
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1500
         X2              =   210
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   5760
         X2              =   4125
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         Height          =   315
         Left            =   5700
         Top             =   870
         Width           =   1935
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   6480
         X2              =   4125
         Y1              =   1170
         Y2              =   1170
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5775
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmCurrencySetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
'MyDDE.SetPermissions = aksess.MayDo("currency setup", aksess.GetID) 'set hak aksess
MyDDE.SetPermissions = aksess.MayDo("Setup Mata Uang") 'set hak aksess

With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmCurrencySetup
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
    .PrepareQuery = "Select * from [Currency Setup]"
    Set DataGrid1(0).DataSource = .ActiveRecordset
End With

'HiasForm Picture1, Me
'HiasFormMantell Picture2, Me
HiasFormManTell Picture2, Me
GridLayout
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
MoveForm Me.hwnd
End Sub

Private Sub Form_Resize()
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MyDDE.CheckRecordPendinged = True Then
   ScanKey vbKeyF5, 0, MyDDE
   If MyDDE.IsSucces = True Then
      Cancel = False
      MyDDE.ClearRecordset
      Set FrmCurrencySetup = Nothing
   Else
      Cancel = True
   End If
Else
   MyDDE.ClearRecordset
   Set FrmCurrencySetup = Nothing
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmCurrencySetup = Nothing
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

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   If MyDDE.ActiveRecordset.Recordcount > 0 Then
      ChkDefault.Value = IIf(MyDDE.GetFieldByName("Functional") = True, 1, 0)
      OptCurrency(0).Value = IIf(MyDDE.GetFieldByName("Cmethod") = "Multiply", 1, 0)
      OptCurrency(1).Value = IIf(MyDDE.GetFieldByName("Cmethod") = "Multiply", 0, 1)
   End If
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
                  MessageBox "Record (" & txtBox(0) & ") Sedang Dipakai Transaksi Lain." & vbCrLf & "Record Tidak Bisa DiHapus.", "Peringatan", msgOkOnly, msgCrtical
                  MyDDE.IsChildMemberReady = False
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbSave:
            If ChkDefault.Value = 1 Then
               SendDataToServer "update [currency setup] set Functional=0"
            End If
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
'Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub PrepareQuery()
On Error GoTo xErr
With MyDDE
    .PrepareAppend = " INSERT INTO [Currency Setup] (CurrID, [Currency Name],[Rate],Functional,source,Cmethod) " & _
                     " VALUES ('" & txtBox(0).Text & "', '" & txtBox(1).Text & "'," & CDbl(txtBox(4)) & "," & ChkDefault.Value & ",'" & txtBox(3).Text & "','" & IIf(OptCurrency(0).Value = True, "Multiply", "Devide") & "')"
                     
    .PrepareUpdate = " UPDATE [Currency Setup] Set [Currency Name] = N'" & txtBox(1) & "',Functional= " & ChkDefault & ",source ='" & txtBox(3).Text & "',Cmethod='" & IIf(OptCurrency(0).Value = True, "Multiply", "Devide") & "' WHERE (CurrID = N'" & ValidString(txtBox(0)) & "')"
    
    .PrepareDelete = " DELETE FROM [Currency Setup] WHERE   (CurrID = N'" & ValidString(txtBox(0)) & "') "
End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub txtBox_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 2 Then ValidNum KeyAscii
End Sub

Private Sub GridLayout()
With DataGrid1(0)
'    Debug.Print .Columns(2).Caption
    .Columns(2).NumberFormat = QtyFormFloat
    .Columns(6).NumberFormat = QtyFormFloat
    .Columns(2).Alignment = dbgRight
    .Columns(3).Alignment = dbgCenter
    .Columns(6).Alignment = dbgRight
    .Columns(5).Alignment = dbgCenter
    .Columns(7).Alignment = dbgCenter
    .Columns(0).width = 1200
    .Columns(3).width = 800
    .Columns(5).width = 800
    .Columns(6).width = 1000
    .Columns(0).Caption = "Mata Uang"
    .Columns(1).Caption = "Negara"
    .Columns(2).Caption = "Kurs"
    .Columns(5).Caption = "Sumber"
    .Columns(6).Caption = "Varian"
    .Columns(7).Caption = "Metode"
End With
End Sub
