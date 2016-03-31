VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmWareHouse 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Warehouse"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8940
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000015&
   Icon            =   "frmWareHouse.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   Tag             =   "Warehouse"
   Begin VB.CommandButton CmdLink 
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   6495
      Picture         =   "frmWareHouse.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1200
      Width           =   330
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4950
      Left            =   0
      ScaleHeight     =   4950
      ScaleWidth      =   8940
      TabIndex        =   4
      Top             =   0
      Width           =   8940
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Locations"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1590
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "Partner"
         Top             =   840
         Width           =   5370
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "WareHouse"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1590
         MaxLength       =   15
         TabIndex        =   5
         Tag             =   "Partner"
         Top             =   150
         Width           =   1935
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "WareHouse Name"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1590
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "Partner"
         Top             =   495
         Width           =   5370
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3105
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Tag             =   "Partner"
         Top             =   1725
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   5477
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "WareHouse"
            Caption         =   "Gudang ID"
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
            DataField       =   "WareHouse Name"
            Caption         =   "Nama Gudang"
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
            DataField       =   "Locations"
            Caption         =   "Lokasi Gudang"
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
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lokasi"
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
         Index           =   2
         Left            =   225
         TabIndex        =   10
         Top             =   885
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Warehouse ID"
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
         Left            =   225
         TabIndex        =   9
         Top             =   195
         Width           =   1170
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
         Left            =   225
         TabIndex        =   8
         Top             =   525
         Width           =   945
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1740
         X2              =   195
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   1740
         X2              =   195
         Y1              =   810
         Y2              =   810
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   1740
         X2              =   195
         Y1              =   1155
         Y2              =   1155
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   1740
         X2              =   195
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Klasifikasi"
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
         Index           =   3
         Left            =   225
         TabIndex        =   7
         Top             =   1230
         Width           =   705
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         DataField       =   "Nama Kelompok Gudang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1590
         TabIndex        =   6
         Tag             =   "Partner"
         Top             =   1185
         Width           =   4920
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5055
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmWareHouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mCall                As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RcPartner                       As New DBQuick
Private mAdd                            As Boolean
Private mVarIndexStr, mVarLastAccount   As String
Private mVarGroupAccount                As String

Private Sub cmdLink_Click(Index As Integer)
OpenPartner Index
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
'HiasForm Picture1, Me
MyDDE.SetPermissions = aksess.MayDo("WareHouse")

HiasFormManTell Picture2, Me
Set mCall = New frmCaller
OpenDB
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
   Else
      Cancel = True
   End If
Else
   MyDDE.ClearRecordset
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mCall = Nothing
Set RcPartner = Nothing
Set frmWareHouse = Nothing
End Sub

Private Sub mCall_BeforeUnload()
On Error Resume Next
MyDDE.SetFocus
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
MyDDE.GetFieldByName("Kode Kelompok Gudang") = mCall.GetFieldByName(0)
MyDDE.GetFieldByName("Nama Kelompok Gudang") = mCall.GetFieldByName(1)
If mAdd = True Then MyDDE.GetFieldByName("NoAccount") = MyAutoIndex
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbAddNew:
            cmdLink(1).Enabled = True
            txtBox(0).SetFocus
            mAdd = True
       Case tmbEdit:
            mAdd = False
            txtBox(0).Enabled = False
            cmdLink(1).Enabled = True
            txtBox(1).SetFocus
       Case tmbSave:

            cmdLink(1).Enabled = False
            mAdd = False
       Case tmbCancel: cmdLink(1).Enabled = False
            mAdd = False
       Case tmbDelete:
            
       Case tmbPrint: CallRPTReport "Tabel Gudang.rpt"
       Case Else: 'mVarDataDc = False
End Select
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
PrepareQuery
Err.Clear
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim mDel As New clsDelete
Select Case AdReasonActiveDb
       Case tmbEdit:
            If MyDDE.ActiveRecordset.Recordcount <> 0 Then
               mVarLastAccount = MyDDE.GetFieldByName("NoAccount")
               'Text2 = mVarLastAccount
               mVarGroupAccount = MyDDE.GetFieldByName("Kode Kelompok Gudang")
            End If
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               If mDel.CekDelete(txtBox(0), reDelMasterGudang) = False Then
                  MyDDE.IsChildMemberReady = True
                  SendDataToServer "Delete From GLAccount where NoAccount='" & mVarLastAccount & "'"
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
            If Label2 = "" Then
               MessageBox "Kelompok persediaan belum dipilih.", "Kelompok persediaan", msgOkOnly
               MyDDE.CancelTrans = True
               Exit Sub
            End If
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               If mAdd = True Then
                  SendDataToServer (" INSERT INTO GLAccount" & _
                                    " (NoAccount, Type, [Group], AccountName, GroupAccount)" & _
                                    " VALUES  (N'" & MyDDE.GetFieldByName("NoAccount") & "', N'" & CariNamaType(37) & "', N'Detail List Account', N'" & ValidString(txtBox(1)) & "', N'" & MyDDE.GetFieldByName("Kode Kelompok Gudang") & "')")
               Else
                  If mVarGroupAccount = MyDDE.GetFieldByName("Kode Kelompok Gudang") Then
                     SendDataToServer " Update GLAccount Set AccountName='" & txtBox(1) & "' where NoAccount='" & mVarLastAccount & "'"
                  Else
                     If SendDataToServer(" Delete From GLAccount where NoAccount='" & mVarLastAccount & "'") = True Then
                        MyDDE.GetFieldByName("NoAccount") = MyAutoIndex
                        SendDataToServer " INSERT INTO GLAccount" & _
                                         " (NoAccount, Type, [Group], AccountName, GroupAccount)" & _
                                         " VALUES  (N'" & MyDDE.GetFieldByName("NoAccount") & "', N'" & CariNamaType(37) & "', N'Detail List Account', N'" & ValidString(txtBox(1)) & "', N'" & MyDDE.GetFieldByName("Kode Kelompok Gudang") & "')"
                        mVarLastAccount = MyDDE.GetFieldByName("NoAccount")
                        mVarGroupAccount = MyDDE.GetFieldByName("Kode Kelompok Gudang")
                     End If
                  End If
               End If
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

Private Sub OpenDB()
With MyDDE
    .EditModeReplace = False
    Set .BindForm = frmWareHouse
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT WareHouse.WareHouse, WareHouse.[WareHouse Name], WareHouse.Locations, WareHouse.NoAccount, WareHouse.GroupAccount AS [Kode Kelompok Gudang], GLAccount.AccountName AS [Nama Kelompok Gudang] FROM         WareHouse INNER JOIN                      GLAccount ON WareHouse.GroupAccount = GLAccount.NoAccount"
End With
End Sub

Private Sub PrepareQuery()
With MyDDE
     .PrepareAppend = " INSERT INTO WareHouse (WareHouse, [WareHouse Name], Locations,GroupAccount,NoAccount) " & _
                      " VALUES (N'" & ValidString(txtBox(0)) & "', N'" & ValidString(txtBox(1)) & "', N'" & ValidString(txtBox(2)) & "','" & MyDDE.GetFieldByName("Kode Kelompok Gudang") & "','" & MyDDE.GetFieldByName("NoAccount") & "')"
                      
     .PrepareUpdate = " UPDATE WareHouse Set NoAccount='" & MyDDE.GetFieldByName("NoAccount") & "',GroupAccount ='" & MyDDE.GetFieldByName("Kode Kelompok Gudang") & "',[WareHouse Name] = N'" & ValidString(txtBox(1)) & "', Locations=N'" & ValidString(txtBox(2)) & "'  WHERE     (WareHouse = N'" & ValidString(txtBox(0)) & "')"
                     
     .PrepareDelete = " DELETE FROM WareHouse WHERE   (WareHouse = N'" & txtBox(0) & "') "
End With
End Sub
Private Sub OpenPartner(ByVal Index As Integer)

On Error GoTo Hell:
Select Case Index
       Case 1:
            RcPartner.DBOpen " SELECT     GLAccount.NoAccount AS [Kode Kelompok Gudang], GLAccount.AccountName AS [Nama Kelompok Gudang] FROM         GLAccount INNER JOIN                       AccType ON GLAccount.Type = AccType.Tipe WHERE     (AccType.ID = 37) AND (GLAccount.[Group] = N'list Account')", CNN, lckLockReadOnly
'       Case 2:
'            RcPartner.DBOpen "SELECT Inventory.NoItem, Inventory.ItemName, Inventory.UOM, Inventory.PPn, MAX([Inventory Tabel].PriceIn) * (Inventory.PPn / 100)  + MAX([Inventory Tabel].PriceIn) * (Inventory.Markup / 100) + MAX([Inventory Tabel].PriceIn) AS Harga, SUM([Inventory Tabel].QTY_IN) AS QTY FROM Inventory LEFT OUTER JOIN [Inventory Tabel] ON Inventory.NoItem = [Inventory Tabel].NoItem WHERE     ([Inventory Tabel].LockFIFO = 0) GROUP BY Inventory.NoItem, Inventory.ItemName, Inventory.PPn, Inventory.Markup, Inventory.UOM HAVING      (SUM([Inventory Tabel].QTY_IN) <> 0)", Cnn, lckLockReadOnly
'            DGPurchase.Columns(6).Visible = False
'            DGPurchase.Columns(7).Visible = True
'
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 1:
            mCall.FromTagActive = "Nama Kelompok Gudang"
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me
Else
   MessageBox "Data Kelompok Gudang Belum Ada.", "Peringatan", msgOkOnly
End If
Exit Sub
Hell:
    Err.Clear
End Sub

Private Function AutoIndexAcc(ByVal GroupAcc As String) As String
Dim Rckode As New DBQuick
Dim mVarTotalDigit As Long
Rckode.DBOpen "SELECT     MAX(RIGHT(NoAccount, 2)) AS MaxNom FROM         GLAccount WHERE     (GroupAccount = N'" & GroupAcc & "') AND ([Group] = N'Detail List Account')", CNN, lckLockReadOnly
With Rckode.DBRecordset
     If .Recordcount <> 0 Then
        mVarTotalDigit = IIf(Not IsNull(.Fields(0)), .Fields(0), 10) + 1
     Else
        mVarTotalDigit = 10
     End If
End With
AutoIndexAcc = Left(GroupAcc, Len(GroupAcc) - 2) & mVarTotalDigit
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

Public Property Get MyAutoIndex() As String
       mVarIndexStr = AutoIndexAcc(MyDDE.GetFieldByName("Kode Kelompok Gudang"))
       MyAutoIndex = mVarIndexStr
End Property

Private Sub GridLayout()
DataGrid1(0).Columns(0).width = 1769.953
DataGrid1(0).Columns(1).width = 3585.26
DataGrid1(0).Columns(2).width = 1934.929
End Sub
