VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmKelompok 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kelompok Persediaan"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmKelompok.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8895
   Tag             =   "Kelompok Stok"
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   10
      Top             =   5055
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5070
      Left            =   0
      ScaleHeight     =   5070
      ScaleWidth      =   8895
      TabIndex        =   11
      Top             =   0
      Width           =   8895
      Begin VB.ComboBox CmbJenis 
         Appearance      =   0  'Flat
         DataField       =   "jenis"
         Height          =   330
         ItemData        =   "frmKelompok.frx":6852
         Left            =   1455
         List            =   "frmKelompok.frx":6860
         TabIndex        =   3
         Tag             =   "Partner"
         Text            =   "CmbJenis"
         Top             =   885
         Width           =   2175
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Group Name"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   1
         Left            =   1455
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "Partner"
         Top             =   540
         Width           =   3045
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "NoGroup"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   0
         Left            =   1455
         MaxLength       =   6
         TabIndex        =   1
         Tag             =   "Partner"
         Top             =   195
         Width           =   1935
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   6030
         Picture         =   "frmKelompok.frx":6885
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1268
         Width           =   330
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmKelompok.frx":D0D7
         Height          =   2895
         Index           =   0
         Left            =   150
         TabIndex        =   9
         Tag             =   "Partner"
         Top             =   2025
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   5106
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         HeadLines       =   2
         RowHeight       =   15
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "NoGroup"
            Caption         =   "NoGroup"
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
            DataField       =   "Group Name"
            Caption         =   "Group Name"
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
            DataField       =   "NoAccount"
            Caption         =   "NoAccount"
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
            DataField       =   "AccountName"
            Caption         =   "AccountName"
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
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
         EndProperty
      End
      Begin VB.Label LBLKodeRek 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "GroupAccount"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   1455
         TabIndex        =   7
         Tag             =   "Partner"
         Top             =   1620
         Width           =   1500
      End
      Begin VB.Label LBLKodeRek 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "NoAccount"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   1455
         TabIndex        =   4
         Tag             =   "Partner"
         Top             =   1260
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelompok"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   15
         Top             =   960
         Width           =   1080
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   1725
         X2              =   150
         Y1              =   1185
         Y2              =   1185
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "GroupAccName"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2985
         TabIndex        =   8
         Tag             =   "Partner"
         Top             =   1620
         Width           =   3045
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group Account"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   14
         Top             =   1695
         Width           =   1065
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   1695
         X2              =   150
         Y1              =   1935
         Y2              =   1935
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   1695
         X2              =   150
         Y1              =   1575
         Y2              =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   13
         Top             =   1335
         Width           =   585
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "AccountName"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2985
         TabIndex        =   5
         Tag             =   "Partner"
         Top             =   1260
         Width           =   3045
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   1725
         X2              =   150
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1725
         X2              =   150
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   0
         Top             =   255
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   12
         Top             =   600
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmKelompok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mCall                As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private mVarIndexStr                            As String
Private mVarLastAccount, mVarGroupAccount       As String
Private mAdd                                    As Boolean
Private RcPartner                       As New DBQuick
Dim strSQL As String

Private Sub Form_Load()
HiasFormManTell Picture2, Me
Set mCall = New frmCaller
OpenDB
GridLayout
'CmbJenis.Clear
'CmbJenis.AddItem "Persediaan"
'CmbJenis.AddItem "Aktiva Tetap"
'CmbJenis.AddItem "Biaya"
End Sub
Private Sub cmdLink_Click(Index As Integer)
Dim TipeAcc As Integer
Select Case UCase(CmbJenis.Text)
    Case "PERSEDIAAN"
        TipeAcc = 37
    Case "AKTIVA TETAP"
        TipeAcc = 5
    Case "BIAYA"
        TipeAcc = 45
End Select
OpenPartner Index, TipeAcc
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
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

'HiasForm Picture1, Me
'CenterForm Picture2, Me
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmKelompok = Nothing
Set mCall = Nothing
Set RcPartner = Nothing
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
    Case tmbAddNew:
         'mVarDataDc = True
         txtBox(0).SetFocus
'         MyDDE.GetFieldByName("NoAccount") = MyAutoIndex
'         MyDDE.GetFieldByName("GroupAccount") = mVarGroupAccount
         cmdLink(1).Enabled = True
    Case tmbEdit:
         txtBox(0).Enabled = False
         cmdLink(1).Enabled = True
         If txtBox(1).Enabled = True Then txtBox(1).SetFocus
    Case tmbDelete:
         
    Case tmbSave:
    '            If mAdd = True Then
    '               SendDataToServer (" INSERT INTO GLAccount" & _
    '                                 " (NoAccount, Type, [Group], AccountName, GroupAccount)" & _
    '                                 " VALUES  (N'" & MyDDE.GetFieldByName("NoAccount") & "', N'" & CariNamaType(37) & "', N'List Account', N'" & ValidString(txtBox(1)) & "', N'" & MyDDE.GetFieldByName("GroupAccount") & "')")
    '            Else
    '               SendDataToServer ("Update GLAccount Set AccountName='" & ValidString(txtBox(1)) & "' where Noaccount='" & MyDDE.GetFieldByName("NoAccount") & "'")
    '            End If
         mAdd = False
    Case tmbPrint:
         CallRPTReport "Tabel Kelompok.rpt"
    Case Else: cmdLink(1).Enabled = False
End Select
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
PrepareQuery
Err.Clear
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   'cboBahan.ListIndex = IIf(Not IsNull(MyDDE.GetFieldByName("Status")), MyDDE.GetFieldByName("Status"), 0)
   'If MyDDE.ActiveRecordset.Recordcount > 0 Then
   'CmbJenis.Text = IIf(IsNull(MyDDE.GetFieldByName("jenis")), "", MyDDE.GetFieldByName("jenis"))
   'End If
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim mDel As New clsDelete
Dim CascaDel As String

Select Case AdReasonActiveDb
       Case tmbAddNew:
            mAdd = True
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
                If mDel.CekDelete(txtBox(0), reDelMasterKelompok) = False Then
                   MyDDE.IsChildMemberReady = True
'                   SendDataToServer ("Delete From GLAccount  where Noaccount='" & MyDDE.GetFieldByName("NoAccount") & "'")
                   PrepareQuery
                Else
                    MyDDE.CancelTrans = True
                    CascaDel = mDel.CekCascadeDeleteTable(txtBox(0), reDelMasterKelompok)
'                    MessageBox "Record (" & txtBox(0) & ") Sedang Dipakai Transaksi Lain." & vbCrLf & "Record Tidak Bisa DiHapus Atau Diedit.", "Peringatan", msgOkOnly
                    MessageBox "Data ' " & txtBox(1).Text & " (" & txtBox(0) & ") '  sedang digunakan pada transaksi " & _
                    UCase(CascaDel) & vbCrLf & "Record tidak bisa dihapus..", "Kontrol Penghapusan Data", msgOkOnly, msgCrtical
                    MyDDE.IsChildMemberReady = False
                End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbEdit:
'            If MyDDE.CheckEmptyControl = False Then
               If mDel.CekDelete(txtBox(0), reDelMasterKelompok) = False Then
                  MyDDE.IsChildMemberReady = True
                  PrepareQuery
                  mAdd = False
'               Else
'                    MyDDE.CancelTrans = True
'                    CascaDel = mDel.CekCascadeDeleteTable(txtBox(0), reDelMasterKelompok)
'                    MessageBox "Data ' " & txtBox(1).Text & " (" & txtBox(0) & ") '  sedang digunakan pada transaksi " & _
'                    UCase(CascaDel) & vbCrLf & "Record tidak bisa dihapus..", "Kontrol Penghapusan Data", msgOkOnly, msgCrtical
'                    MyDDE.IsChildMemberReady = False
                    
'                  MessageBox "Record (" & txtBox(0) & ") Sedang Dipakai Transaksi Lain." & vbCrLf & "Record Tidak Bisa DiHapus Atau Diedit.", "Peringatan", msgOkOnly
                    
               End If
           ' Else
           '    MyDDE.IsChildMemberReady = False
           ' End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               PrepareQuery
'               MyDDE.GetFieldByName("Status") = cboBahan.ListIndex
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
Set mDel = Nothing
End Sub

Private Sub OpenDB()
With MyDDE
    .EditModeReplace = False
    Set .BindForm = frmKelompok
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
'    .PrepareQuery = "SELECT [Inventory Group].NoGroup, [Inventory Group].[Group Name], [Inventory Group].Status, " & _
                    " [Inventory Group].NoAccount, [Inventory Group].GroupAccount, [Inventory Group].jenis, GLAccount.AccountName " & _
                    " FROM [Inventory Group] LEFT OUTER JOIN GLAccount ON [Inventory Group].NoAccount = GLAccount.NoAccount"
                    
    strSQL = "SELECT [Inventory Group].NoGroup, [Inventory Group].[Group Name], [Inventory Group].Status, [Inventory Group].NoAccount, " & _
            " [Inventory Group].GroupAccount, [Inventory Group].jenis, GLAccount_1.AccountName, GLAccount.AccountName AS GroupAccName " & _
            " FROM [Inventory Group] LEFT OUTER JOIN GLAccount ON [Inventory Group].GroupAccount = GLAccount.NoAccount " & _
            " LEFT OUTER JOIN GLAccount AS GLAccount_1 ON [Inventory Group].NoAccount = GLAccount_1.NoAccount"
    .PrepareQuery = strSQL
'    Debug.Print .PrepareQuery
End With
End Sub

Private Sub PrepareQuery()
On Error GoTo xErr
With MyDDE
    .PrepareAppend = " INSERT INTO [Inventory Group] (NoGroup, [Group Name],Status,NoAccount,GroupAccount, jenis) " & _
                     " VALUES (N'" & ValidString(txtBox(0)) & "', N'" & ValidString(txtBox(1)) & "',0,'" & _
                     MyDDE.GetFieldByName("NoAccount") & "','" & MyDDE.GetFieldByName("GroupAccount") & "','" & CmbJenis.Text & "')"
                     
    .PrepareUpdate = " UPDATE [Inventory Group] Set Status=0,[Group Name] = N'" & ValidString(txtBox(1)) & "', " & _
                " NoAccount = " & FNumText(MyDDE.GetFieldByName("NoAccount")) & ",GroupAccount=" & FNumText(MyDDE.GetFieldByName("GroupAccount")) & ", jenis = " & FText(CmbJenis.Text) & " WHERE (NoGroup = N'" & ValidString(txtBox(0)) & "')"
                     
    .PrepareDelete = " DELETE FROM [Inventory Group] WHERE   (NoGroup = N'" & ValidString(txtBox(0)) & "') "
'MessageBox .PrepareUpdate
End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

'Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
'End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Public Property Get MyAutoIndex() As String
       'AccountGroup
       mVarIndexStr = AutoIndexAcc
       MyAutoIndex = mVarIndexStr
End Property

Private Sub AccountGroup(TipeAcc As Integer)
Dim Rc As New DBQuick
Rc.DBOpen " SELECT GLAccount.NoAccount AS [Kode Kelompok Barang], GLAccount.AccountName AS [Nama Kelompok Barang], GLAccount.GroupAccount" & _
          " FROM GLAccount INNER JOIN AccType ON GLAccount.Type = AccType.Tipe WHERE     (GLAccount.[Group] = N'List Account') AND (AccType.ID = " & TipeAcc & ")", CNN, lckLockReadOnly
With Rc
     If .Recordcount <> 0 Then
        mVarLastAccount = IIf(Not IsNull(.Fields(0)), .Fields(0), "")
        mVarGroupAccount = IIf(Not IsNull(.Fields(2)), .Fields(2), "")
     End If
End With
End Sub

Private Function AutoIndexAcc() As String
Dim Rckode As New DBQuick
Dim mVarTotalDigit As Long
'AccountGroup
Rckode.DBOpen "SELECT   MAX(SUBSTRING(NoAccount, 7, 2)) AS MaxNom FROM GLAccount " & _
        " WHERE (GroupAccount = N'" & mVarGroupAccount & "') AND ([Group] = N'List Account')", CNN, lckLockReadOnly
With Rckode.DBRecordset
     If .Recordcount <> 0 Then
        mVarTotalDigit = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + 1
     Else
        mVarTotalDigit = 1
     End If
End With
AutoIndexAcc = Left(mVarGroupAccount, 6) & mVarTotalDigit & KirimNull(2)
End Function

Private Function CariNamaType(ByVal Params As Long) As String
Dim RcAkum As New DBQuick
RcAkum.DBOpen "SELECT AccType.tipe FROM AccType where (AccType.ID = " & Params & ")", CNN, lckLockReadOnly
With RcAkum.DBRecordset
     If .Recordcount <> 0 Then
        CariNamaType = IIf(Not IsNull(.Fields(0)), .Fields(0), "")
     End If
End With
End Function

Private Function IsAccDelete(ByVal Params As String) As Boolean
Dim Rc As New DBQuick
Rc.DBOpen "SELECT     NoAccount, GroupAccount FROM         GLAccount WHERE     (GroupAccount = N'" & Params & "')", CNN, lckLockReadOnly
If Rc.Recordcount <> 0 Then IsAccDelete = True
End Function

Private Sub GridLayout()
With DataGrid1(0)
    .Columns(0).width = 800
    .Columns(1).width = 3200
    .Columns(3).width = 2760
    .Columns(0).Caption = "Kode"
    .Columns(1).Caption = "Keterangan"
    .Columns(2).Caption = "Kode Rekening"
    .Columns(3).Caption = "Nama Rekening"
    
End With
End Sub
Private Sub mCall_BeforeUnload()
On Error Resume Next
MyDDE.SetFocus
End Sub
Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
MyDDE.GetFieldByName("NoAccount") = mCall.GetFieldByName(0)
MyDDE.GetFieldByName("AccountName") = mCall.GetFieldByName(1)
MyDDE.GetFieldByName("GroupAccount") = mCall.GetFieldByName(2)
MyDDE.GetFieldByName("GroupAccName") = mCall.GetFieldByName(3)
'If mAdd = True Then MyDDE.GetFieldByName("NoAccount") = MyAutoIndex
End Sub
Private Sub OpenPartner(ByVal Index As Integer, TipeAcc As Integer)

On Error GoTo Hell:
Select Case Index
    Case 1:
        If TipeAcc = 5 Then
            strSQL = "SELECT  GLAccount.NoAccount AS [Kode Kelompok Gudang], GLAccount.AccountName AS [Nama Kelompok Gudang], " & _
                " GLAccount.GroupAccount, GLAccount_1.AccountName AS GroupAccName " & _
                " FROM GLAccount INNER JOIN AccType ON GLAccount.Type = AccType.Tipe INNER JOIN " & _
                " GLAccount AS GLAccount_1 ON GLAccount.GroupAccount = GLAccount_1.NoAccount " & _
                " WHERE (AccType.ID IN (5, 6, 7)) AND (GLAccount.[Group] = N'list Account')"
'            Debug.Print strSQL
            RcPartner.DBOpen strSQL, CNN, lckLockReadOnly
            
'         RcPartner.DBOpen " SELECT GLAccount.NoAccount AS [Kode Kelompok Gudang], GLAccount.AccountName AS [Nama Kelompok Gudang] " & _
         " FROM GLAccount INNER JOIN AccType ON GLAccount.Type = AccType.Tipe " & _
         " WHERE (AccType.ID IN (5, 6, 7)) AND (GLAccount.[Group] = N'list Account')", CNN, lckLockReadOnly
         
        Else
'            RcPartner.DBOpen " SELECT GLAccount.NoAccount AS [Kode Kelompok Gudang], GLAccount.AccountName AS [Nama Kelompok Gudang] " & _
            " FROM GLAccount INNER JOIN AccType ON GLAccount.Type = AccType.Tipe " & _
            " WHERE (AccType.ID = " & TipeAcc & ") AND (GLAccount.[Group] = N'list Account')", CNN, lckLockReadOnly
            
            strSQL = "SELECT  GLAccount.NoAccount AS [Kode Kelompok Gudang], GLAccount.AccountName AS [Nama Kelompok Gudang], " & _
                " GLAccount.GroupAccount, GLAccount_1.AccountName AS GroupAccName " & _
                " FROM GLAccount INNER JOIN AccType ON GLAccount.Type = AccType.Tipe INNER JOIN " & _
                " GLAccount AS GLAccount_1 ON GLAccount.GroupAccount = GLAccount_1.NoAccount " & _
                " WHERE (AccType.ID = " & TipeAcc & ") AND (GLAccount.[Group] = N'list Account')"
            RcPartner.DBOpen strSQL, CNN, lckLockReadOnly
        End If
End Select
If RcPartner.Recordcount <> 0 Then
    Select Case Index
        Case 1:
            Select Case TipeAcc
                Case 37: mCall.FromTagActive = "Kode Rekening Persediaan"
                Case 45: mCall.FromTagActive = "Kode Rekening Biaya"
                Case 5: mCall.FromTagActive = "Kode Rekening Aktiva Tetap"
            End Select
        Case Else
    End Select
    Set mCall.FormData = RcPartner.DBRecordset
    mCall.LookUp Me
Else
   MessageBox "Data kode rekening Belum Ada.", "Peringatan", msgOkOnly, msgCrtical
End If

Exit Sub
Hell:
    MessageBox Err.Description, frmKelompok.Caption, msgOkOnly, msgExclamation
End Sub


