VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{EB0E2EAE-5969-4167-B57F-56BCD8266DF2}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmBankPartner 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bank Partner"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   9990
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBankPartner.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   Tag             =   "Bank Partner"
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   4935
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   0
      ScaleHeight     =   4995
      ScaleWidth      =   9990
      TabIndex        =   8
      Top             =   0
      Width           =   9990
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   " Pilih Kategori "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   120
         TabIndex        =   13
         Top             =   90
         Width           =   3855
         Begin VB.OptionButton Option1 
            BackColor       =   &H00EAAF6F&
            Caption         =   "Customer"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   390
            TabIndex        =   15
            Top             =   255
            Value           =   -1  'True
            Width           =   1395
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00EAAF6F&
            Caption         =   "Supplier"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   1950
            TabIndex        =   14
            Top             =   255
            Width           =   1395
         End
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Default"
         DataField       =   "Default"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   210
         Left            =   4230
         TabIndex        =   5
         Tag             =   "BANK"
         Top             =   1785
         Width           =   1590
      End
      Begin VB.TextBox txtBox 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "Address"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   2
         Left            =   5610
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "BANK"
         Top             =   1050
         Width           =   4275
      End
      Begin VB.TextBox txtBox 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "Bank Name"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   1
         Left            =   5610
         MaxLength       =   25
         TabIndex        =   2
         Tag             =   "BANK"
         Top             =   720
         Width           =   4275
      End
      Begin VB.TextBox txtBox 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "Account"
         Height          =   315
         Index           =   0
         Left            =   5610
         MaxLength       =   25
         TabIndex        =   1
         Tag             =   "BANK"
         Top             =   390
         Width           =   4275
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmBankPartner.frx":6852
         Height          =   2655
         Index           =   0
         Left            =   4035
         TabIndex        =   6
         Tag             =   "BANK"
         Top             =   2175
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   4683
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Account"
            Caption         =   "No. Rekening"
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
            DataField       =   "Bank Name"
            Caption         =   "Nama Bank"
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
            DataField       =   "Address"
            Caption         =   "Alamat"
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
         BeginProperty Column03 
            DataField       =   "Currency"
            Caption         =   "Mata Uang"
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
         BeginProperty Column04 
            DataField       =   "Default"
            Caption         =   "Default"
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
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DGGrid 
         Height          =   4065
         Left            =   120
         TabIndex        =   7
         Top             =   765
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   7170
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Appearance      =   0
         HeadLines       =   2
         RowHeight       =   15
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
            DataField       =   "PartnerID"
            Caption         =   "Partner ID"
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
            DataField       =   "CompanyName"
            Caption         =   "Nama Perusahaan"
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
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboCurr 
         DataField       =   "Currency"
         DataSource      =   "Adodc1"
         Height          =   330
         Left            =   5610
         TabIndex        =   4
         Tag             =   "BANK"
         Top             =   1380
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         ListField       =   "Currency Name"
         BoundColumn     =   "CurrID"
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mata Uang"
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
         Index           =   3
         Left            =   4245
         TabIndex        =   12
         Top             =   1455
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
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
         Left            =   4245
         TabIndex        =   11
         Top             =   1095
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Bank"
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
         Left            =   4245
         TabIndex        =   10
         Top             =   765
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Rekening"
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
         Left            =   4245
         TabIndex        =   9
         Top             =   420
         Width           =   960
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   4215
         X2              =   5670
         Y1              =   690
         Y2              =   690
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   4215
         X2              =   5670
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   4200
         X2              =   5655
         Y1              =   1350
         Y2              =   1350
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   4200
         X2              =   5655
         Y1              =   1695
         Y2              =   1695
      End
   End
End
Attribute VB_Name = "frmBankPartner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents RcPartner As Recordset
Attribute RcPartner.VB_VarHelpID = -1
Dim Rc As New Recordset
Dim mVarNoAccount As String
'Dim mVarDefaultCurr, mAdd, mRubah As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

'Private Sub cboCurr_Click(Area As Integer)
'mVarMataUang = cboCurr.BoundText
'End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
End Sub

Private Sub cboCurr_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub cmdOk_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub Form_Load()
On Error GoTo 1
MyDDE.SetPermissions = aksess.MayDo("Bank Partner")

GridLayout
OpenPartner "Customer"
CloseDB Rc
Set Rc = New Recordset
Rc.CursorLocation = adUseClient
Rc.Open "SELECT * FROM [Currency Setup] ORDER BY CurrID", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
Set cboCurr.RowSource = Rc
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
Check1.BackColor = Picture2.BackColor
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
Option1(0).BackColor = Picture2.BackColor
Option1(1).BackColor = Picture2.BackColor
Check1.BackColor = Picture2.BackColor
Exit Sub
1:
MessageBox Err.Description, "frmbankpartner:form_load" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Form_Resize()

Err.Clear
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 1
Select Case AdReasonActiveDb
       Case tmbEdit, tmbAddNew:
            If RcPartner.Recordcount <> 0 Then
               mVarNoAccount = txtBox(0)
               MyDDE.CancelTrans = False
               MyDDE.GetFieldByName("Default") = 1
               If txtBox(0).Enabled = True Then txtBox(0).SetFocus
               DGGrid.Enabled = False
            Else
               MessageBox "Data Partner Belum Ada.", "Peringatan", msgOkOnly, msgCrtical
               MyDDE.CancelTrans = True
            End If
       Case tmbSave, tmbCancel: DGGrid.Enabled = True
End Select
Exit Sub
1:
MessageBox Err.Description, "frmbankpartner:mydde_afterprepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 2
Dim mDel As New clsDelete
Select Case AdReasonActiveDb
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               If mDel.CekDelete(txtBox(0), reDelMasterBank) = False Then
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
       Case tmbEdit, tmbAddNew:
            If RcPartner.Recordcount <> 0 Then
               MyDDE.CancelTrans = False

            Else
               MessageBox "Data Partner Belum Ada.", "Peringatan", msgOkOnly, msgCrtical
               MyDDE.CancelTrans = True
            End If
       Case tmbSave:
            If RcPartner.Recordcount <> 0 Then
                If MyDDE.CheckEmptyControl = False Then
                   If CekRekening = False Then
                      MyDDE.IsChildMemberReady = True
                      PrepareQuery
                   Else
                      MessageBox "Nomer Rekening Sudah ada.", "Peringatan", msgOkOnly, msgCrtical
                      MyDDE.IsChildMemberReady = False
                   End If
                Else
                   MyDDE.IsChildMemberReady = False
                End If
            Else
                MessageBox "Data Partner Belum Ada.", "Peringatan", msgOkOnly, msgCrtical
                
            End If
End Select
Set mDel = Nothing
Exit Sub
2:
MessageBox Err.Description, "frmbankpartner:mydde_afterprepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub
Private Sub Option1_Click(Index As Integer)
If Index = 0 Then OpenPartner "Customer" Else OpenPartner "Supplier"
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

Private Sub rcPartner_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo 1
With pRecordset
     If .Recordcount <> 0 Then
        OpenBank IIf(Not IsNull(.Fields("PartnerID")), .Fields("PartnerID"), "XXXX")
     End If
End With
Exit Sub
1:
MessageBox Err.Description, "frmbankpartner:rcpartner_movecomplete" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Function ChekKosong() As Boolean
On Error GoTo 2

Dim I As Integer
For I = 0 To 2
    If txtBox(I).Text = "" Then
       ChekKosong = True
       Exit For
    End If
Next I
If cboCurr = "" Then ChekKosong = True
Exit Function
2:
MessageBox Err.Description, "frmbankpartner:checkkosong" & Err.Number, msgOkOnly, msgExclamation
End Function

Private Sub DefCurr()
On Error GoTo 3
Dim RcCurr As New Recordset
RcCurr.CursorLocation = adUseClient
RcCurr.Open "SELECT     [Currency Name] FROM         [Currency Setup] WHERE     (CurrID = N'" & cboCurr.BoundText & "')", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
With RcCurr
     If .Recordcount <> 0 Then
        cboCurr.Text = IIf(Not IsNull(.Fields(0)), .Fields(0), "")
     End If
End With
CloseDB RcCurr
Err.Clear
Exit Sub
3:
MessageBox Err.Description, "frmbankpartner:defcurr" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub OpenPartner(ByVal Tipical As String)
On Error GoTo 6
CloseDB RcPartner
Set RcPartner = New Recordset
RcPartner.CursorLocation = adUseClient
RcPartner.Open "SELECT     PartnerID, CompanyName FROM PartnerDB WHERE     (PartnerType = N'" & Tipical & "') ORDER BY PartnerID", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
Set RcPartner.ActiveConnection = Nothing
With RcPartner
     Set DGGrid.DataSource = RcPartner
End With
Exit Sub
6:
MessageBox Err.Description, "frmbankpartner:openpartner" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub OpenBank(ByVal NoPartner As String)
On Error GoTo 5
With MyDDE
    .EditModeReplace = False
    Set .BindForm = frmBankPartner
    .BindFormTAG = "BANK"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT Account, [Bank Name], Address, Currency, [Default] FROM [Bank Partner] WHERE (PartnerID = N'" & NoPartner & "')"
End With
Exit Sub
5:
MessageBox Err.Description, "frmbankpartner:openbank" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub PrepareQuery()
On Error GoTo xErr
With MyDDE
    .PrepareAppend = " INSERT INTO [Bank Partner]" & _
                     " (Account, [Bank Name], Address, Currency, [Default], PartnerID)" & _
                     " VALUES (N'" & ValidString(txtBox(0)) & "', N'" & ValidString(txtBox(1)) & "', N'" & ValidString(txtBox(2)) & "', N'" & cboCurr.BoundText & "', " & BoolToInt(MyDDE.GetFieldByName("Default")) & ", N'" & RcPartner.Fields(0) & "')"
                     
    .PrepareUpdate = " UPDATE [Bank Partner]" & _
                     " Set Account =N'" & ValidString(txtBox(0)) & "' ,[Bank Name] = N'" & ValidString(txtBox(1)) & "', Address = N'" & ValidString(txtBox(2)) & "', Currency = N'" & cboCurr.BoundText & "', [Default] = " & BoolToInt(MyDDE.GetFieldByName("Default")) & _
                     " WHERE (PartnerBank = N'" & IdxBank & "') "
                     
    .PrepareDelete = " DELETE FROM [Bank Partner] WHERE (PartnerBank=N'" & IdxBank & "')"
End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear
End Sub

Private Function IdxBank() As String
On Error GoTo 4
Dim RcId As New Recordset
RcId.CursorLocation = adUseClient
RcId.Open "SELECT PartnerBank FROM [Bank Partner] WHERE     (PartnerID = N'" & RcPartner.Fields(0) & "') AND (Account = N'" & mVarNoAccount & "')", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
Set RcId.ActiveConnection = Nothing
IdxBank = "XXX"
With RcId
     If .Recordcount <> 0 Then
        IdxBank = .Fields(0)
     End If
End With
CloseDB RcId
Exit Function
4:
MessageBox Err.Description, "frmbankpartner:idxbank" & Err.Number, msgOkOnly, msgExclamation
End Function

Private Function CekRekening() As Boolean
On Error GoTo 1
Dim Rcb As New Recordset
Rcb.CursorLocation = adUseClient
Rcb.Open "SELECT PartnerID, Account FROM [Bank Partner] WHERE     (PartnerID = N'" & RcPartner.Fields(0) & "') AND (Account = N'" & txtBox(0) & "')", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
With Rcb
     If .Recordcount <> 0 Then
        CekRekening = True
     Else
        CekRekening = False
     End If
End With
CloseDB Rcb
Exit Function
1:
MessageBox Err.Description, "frmbankpartner:cekrekening" & Err.Number, msgOkOnly, msgExclamation
End Function

Private Sub GridLayout()
DGGrid.Columns(0).width = 1214.929
DGGrid.Columns(1).width = 2085.166
DataGrid1(0).Columns(0).width = 1739.906
DataGrid1(0).Columns(1).width = 3225.26
DataGrid1(0).Columns(2).width = 1844.787
DataGrid1(0).Columns(3).width = 2115.213
DataGrid1(0).Columns(4).width = 2835.213
End Sub
