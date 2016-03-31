VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmCurrencyMaint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nilai Tukar Mata Uang"
   ClientHeight    =   7095
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
   Icon            =   "FrmCurrencyMaint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   Tag             =   "Exchange Rate Maintenance"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   6525
      Left            =   0
      ScaleHeight     =   6525
      ScaleWidth      =   8865
      TabIndex        =   8
      Top             =   0
      Width           =   8865
      Begin MSComCtl2.DTPicker DTPExcDate 
         DataField       =   "ExpDate"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   1
         Left            =   5565
         TabIndex        =   6
         Tag             =   "CURR"
         Top             =   1155
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   71630851
         CurrentDate     =   38773
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "CurrID"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   0
         Left            =   1410
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   1
         Tag             =   "CURR"
         Top             =   120
         Width           =   1935
      End
      Begin VB.TextBox txtBox 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "Rate"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00;(#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         DataSource      =   "MyDDE"
         Height          =   320
         Index           =   2
         Left            =   1410
         MaxLength       =   50
         TabIndex        =   5
         Tag             =   "CURR"
         Top             =   1155
         Width           =   1935
      End
      Begin VB.CommandButton CmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   3345
         Picture         =   "FrmCurrencyMaint.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "CURR"
         Top             =   128
         Width           =   330
      End
      Begin MSComCtl2.DTPicker DTPExcDate 
         DataField       =   "ExcDate"
         DataSource      =   "MyDDE"
         Height          =   320
         Index           =   0
         Left            =   1410
         TabIndex        =   3
         Tag             =   "CURR"
         Top             =   465
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   71630851
         CurrentDate     =   38773
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmCurrencyMaint.frx":6BDC
         Height          =   4830
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   1590
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   8520
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "IdxCurr"
            Caption         =   "IdxCurr"
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
            DataField       =   "CurrID"
            Caption         =   "Mata Uang"
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
            DataField       =   "ExcDate"
            Caption         =   "Tanggal"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd MMMM yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "ExcTime"
            Caption         =   "Waktu"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "h:mm:ss AMPM"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   4
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Rate"
            Caption         =   "Kurs"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "ExpDate"
            Caption         =   "Tanggal Kadaluarsa"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dddd dd MMMM yyyy"
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
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
            EndProperty
            BeginProperty Column05 
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPExcDate 
         DataField       =   "ExcTime"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   2
         Left            =   1410
         TabIndex        =   4
         Tag             =   "CURR"
         Top             =   810
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   71630850
         CurrentDate     =   38773
      End
      Begin VB.Label LBLCurrency 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Negara"
         DataField       =   "Negara"
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
         Height          =   270
         Index           =   4
         Left            =   4080
         TabIndex        =   18
         Tag             =   "CURR"
         Top             =   150
         Width           =   3045
      End
      Begin VB.Label LBLCurrency 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Description"
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
         Height          =   270
         Index           =   3
         Left            =   5535
         TabIndex        =   11
         Tag             =   "CURR"
         Top             =   840
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.Label LBLCurrency 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Description"
         DataField       =   "Address"
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
         Height          =   270
         Index           =   2
         Left            =   5535
         TabIndex        =   10
         Top             =   495
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Kurs"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   525
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mata Uang"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   180
         Width           =   780
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2000
         X2              =   210
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   1
         X1              =   2000
         X2              =   210
         Y1              =   770
         Y2              =   770
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   2
         X1              =   2000
         X2              =   195
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kurs"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   1215
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   870
         Width           =   465
      End
      Begin VB.Line Line2 
         X1              =   195
         X2              =   1905
         Y1              =   1455
         Y2              =   1455
      End
      Begin VB.Label LBLCurrency 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         DataField       =   "Address"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   4080
         TabIndex        =   13
         Tag             =   "PO"
         Top             =   525
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label LBLCurrency 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sumber Kurs"
         DataField       =   "Address"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   4080
         TabIndex        =   12
         Tag             =   "PO"
         Top             =   870
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Line Line1 
         Index           =   3
         Visible         =   0   'False
         X1              =   6120
         X2              =   4050
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line Line1 
         Index           =   4
         Visible         =   0   'False
         X1              =   6120
         X2              =   4050
         Y1              =   1095
         Y2              =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   6
         X1              =   6120
         X2              =   4050
         Y1              =   1455
         Y2              =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Kadaluarsa"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   5
         Left            =   4080
         TabIndex        =   9
         Top             =   1230
         Width           =   1410
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      CausesValidation=   0   'False
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   6525
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   1005
      BindFormTAG     =   "CURR"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmCurrencyMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RsCurrency As New DBQuick
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Dim strSQL As String

Private Sub GridLayout()
With DataGrid1(0)
    .Columns(4).NumberFormat = QtyFormFloat
    .Columns(5).width = 1960
    .Columns(2).NumberFormat = ShortDateForm
    .Columns(5).NumberFormat = ShortDateForm
End With
End Sub
Private Sub cmdLink_Click(Index As Integer)
   OpenPartner Index
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()

'MyDDE.SetPermissions = aksess.MayDo("Exchange Rate Maintenan", aksess.GetID)
MyDDE.SetPermissions = aksess.MayDo("Nilai Tukar Mata Uang")

With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmCurrencyMaint
    .BindFormTAG = "CURR"
    Set .ActiveConnection = CNN
    strSQL = "SELECT  [Currency Maint].IdxCurr, [Currency Maint].CurrID, [Currency Maint].ExcDate, [Currency Maint].ExcTime, " & _
        " [Currency Maint].Rate, [Currency Maint].ExpDate, [Currency Setup].[Currency Name] AS Negara " & _
        " FROM [Currency Setup] INNER JOIN [Currency Maint] ON [Currency Setup].CurrID = [Currency Maint].CurrID " & _
        " ORDER BY [Currency Maint].ExcDate DESC"

    .PrepareQuery = strSQL
    Set DataGrid1(0).DataSource = MyDDE.ActiveRecordset
End With
'HiasForm Picture1, Me
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
      Set FrmCurrencyMaint = Nothing
   Else
      Cancel = True
   End If
Else
   MyDDE.ClearRecordset
   Set FrmCurrencyMaint = Nothing
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmCurrencyMaint = Nothing
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
   If TagForm = "Master Currency" Then
      MyDDE.GetFieldByName("CurrID") = mCall.GetFieldByName("Currency")
   End If
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 1
Dim rsIndex As New DBQuick
Select Case AdReasonActiveDb
       Case tmbAddNew:
            'mVarDataDc = True
            cmdLink(0).Enabled = True
            cmdLink(0).SetFocus
            DTPExcDate(0).Value = Now
            DTPExcDate(1).Value = Now
            DTPExcDate(2).Value = Now
            MyDDE.GetFieldByName("ExcDate") = Now
            MyDDE.GetFieldByName("ExcTime") = Now
            MyDDE.GetFieldByName("ExpDate") = Now
            
            rsIndex.DBOpen "select newID()", CNN, lckLockReadOnly
            MyDDE.GetFieldByName("IdxCurr") = rsIndex.DBRecordset.Fields(0)
            Set rsIndex = Nothing
            
       Case tmbEdit:
            txtBox(0).Enabled = False
            'mVarDataDc = True
            DTPExcDate(0).SetFocus
            cmdLink(0).Enabled = True
       Case tmbPrint:
            CallRPTReport "Maint Mata Uang.rpt"
       Case tmbCancel:
            cmdLink(0).Enabled = False 'mVarDataDc = False
       Case tmbSave:
            SaveToMasterCurrency
            cmdLink(0).Enabled = False 'mVarDataDc = False
End Select
Exit Sub
1:
MessageBox Err.Description, "frmcurrencymaint:mydde_afterprepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub SaveToMasterCurrency()
On Error GoTo xErr
   Dim rsCurrSelect As New DBQuick
   rsCurrSelect.DBOpen "select CurrID from [currency Maint] where CurrID = '" & MyDDE.GetFieldByName("CurrID") & "' order by Excdate desc ,ExcTime desc", CNN, lckLockReadOnly
   SendDataToServer "update [Currency Setup] set Rate = " & MyDDE.GetFieldByName("Rate") & " where CurrID='" & MyDDE.GetFieldByName("CurrID") & "'"
   Set rsCurrSelect = Nothing
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

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
MessageBox Err.Description, "frmcurrencymaint:mydde_prepareactivedb" & Err.Number, msgOkOnly, msgExclamation
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
    .PrepareAppend = " INSERT INTO [Currency Maint] (idxCurr,CurrID, ExcDate,ExcTime,Rate,ExpDate) " & _
                     " VALUES ('" & MyDDE.GetFieldByName("idxCurr") & "','" & ValidString(txtBox(0).Text) & "', '" & Format(DTPExcDate(0).Value, "yyyy-MM-dd") & "','" & Format(DTPExcDate(2).Value, "yyyy-MM-dd hh:mm:ss") & "'," & ValidString(txtBox(2).Text) & ",'" & Format(DTPExcDate(1).Value, "yyyy-MM-dd") & "')"
                     
    .PrepareUpdate = " UPDATE [Currency Maint] Set CurrID = '" & txtBox(0).Text & "' , Rate =" & txtBox(2).Text & ", ExcDate='" & Format(DTPExcDate(0).Value, "yyyy-MM-dd") & "', ExcTime ='" & Format(DTPExcDate(2).Value, "yyyy-MM-dd hh:mm:ss") & "', ExpDate ='" & Format(DTPExcDate(1).Value, "yyyy-MM-dd") & "' WHERE (CurrID = '" & ValidString(txtBox(0)) & "')"
    
    .PrepareDelete = " DELETE FROM [Currency Maint] WHERE   (IdxCurr = '" & MyDDE.GetFieldByName("IdxCurr") & "')"
End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub


Private Sub txtBox_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 2 Then ValidNum KeyAscii
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell
Set mCall = New frmCaller
Select Case Index
       Case 0:
            RsCurrency.DBOpen " SELECT [Currency Setup].CurrID As Currency, " & _
            " [Currency Setup].[Currency Name] AS Description FROM [Currency Setup]", _
            CNN, lckLockReadOnly
End Select

If RsCurrency.Recordcount <> 0 Then
    If Index = 0 Then
       mCall.FromTagActive = "Master Currency"
       'mCall.txtCari = txtBox(0)
    End If
    Set mCall.FormData = RsCurrency.DBRecordset
    mCall.LookUp Me
'    cboVoucher.Enabled = False
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgCrtical
End If

Exit Sub
Hell:
    'messagebox Err.Description
    Err.Clear
End Sub
