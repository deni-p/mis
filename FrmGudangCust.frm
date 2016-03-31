VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{EB0E2EAE-5969-4167-B57F-56BCD8266DF2}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmGudangCust 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Warehouse"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmGudangCust.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   Tag             =   "Customer Warehouse"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5010
      Left            =   0
      ScaleHeight     =   5010
      ScaleWidth      =   9240
      TabIndex        =   8
      Top             =   0
      Width           =   9240
      Begin MSDataGridLib.DataGrid DGGrid 
         Height          =   4530
         Left            =   120
         TabIndex        =   7
         Top             =   405
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   7990
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Appearance      =   0
         BorderStyle     =   0
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
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2910
         Index           =   0
         Left            =   4080
         TabIndex        =   6
         Tag             =   "BANK"
         Top             =   2040
         Width           =   4890
         _ExtentX        =   8625
         _ExtentY        =   5133
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   0
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
            DataField       =   "GDG ID"
            Caption         =   "GDG ID"
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
            DataField       =   "Nama Gudang"
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
            DataField       =   "Alamat"
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
            DataField       =   "Kota"
            Caption         =   "Kota"
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
            DataField       =   "Telp"
            Caption         =   "Telp"
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
            BeginProperty Column04 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Telp"
         Height          =   330
         Index           =   3
         Left            =   5550
         MaxLength       =   15
         TabIndex        =   5
         Tag             =   "BANK"
         Top             =   1560
         Width           =   3420
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "GDG ID"
         Height          =   330
         Index           =   0
         Left            =   5550
         MaxLength       =   15
         TabIndex        =   1
         Tag             =   "BANK"
         Top             =   105
         Width           =   3420
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Nama Gudang"
         Height          =   330
         Index           =   1
         Left            =   5550
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "BANK"
         Top             =   480
         Width           =   3420
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Alamat"
         Height          =   330
         Index           =   2
         Left            =   5550
         MaxLength       =   100
         TabIndex        =   3
         Tag             =   "BANK"
         Top             =   840
         Width           =   3420
      End
      Begin MSDataListLib.DataCombo cboCurr 
         DataField       =   "RG"
         Height          =   330
         Left            =   5550
         TabIndex        =   4
         Tag             =   "BANK"
         Top             =   1200
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         ListField       =   "RG Name"
         BoundColumn     =   "RG"
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City"
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
         Left            =   4200
         TabIndex        =   14
         Top             =   1268
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Left            =   4200
         TabIndex        =   13
         Top             =   908
         Width           =   585
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
         Left            =   4200
         TabIndex        =   12
         Top             =   548
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Warehouse ID"
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
         Left            =   4200
         TabIndex        =   11
         Top             =   173
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telphone"
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
         Index           =   4
         Left            =   4200
         TabIndex        =   10
         Top             =   1628
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer List"
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
         Index           =   5
         Left            =   135
         TabIndex        =   9
         Top             =   150
         Width           =   1185
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   4185
         X2              =   5685
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   4185
         X2              =   5685
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   4185
         X2              =   5685
         Y1              =   1155
         Y2              =   1155
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   4185
         X2              =   5685
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   4185
         X2              =   5685
         Y1              =   1875
         Y2              =   1875
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5010
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmGudangCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents RcPartner As Recordset
Attribute RcPartner.VB_VarHelpID = -1
Dim Rc As New DBQuick
Dim mVarNoAccount As String

Private Sub cboCurr_GotFocus()
Rc.DBOpen "SELECT RG, [RG Name] FROM Regional ORDER BY RG", CNN, lckLockReadOnly
Set cboCurr.RowSource = Rc.DBRecordset
cboCurr.BoundColumn = "RG"
End Sub

Private Sub cboCurr_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
MyDDE.SetPermissions = aksess.MayDo("Customer Warehouse")
GridLayout
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
OpenPartner "Customer"
Rc.DBOpen "SELECT RG, [RG Name] FROM Regional ORDER BY RG", CNN, lckLockReadOnly
Set cboCurr.RowSource = Rc.DBRecordset
End Sub

Private Sub Form_Resize()

'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
Err.Clear
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit:
            If RcPartner.Recordcount <> 0 Then
               txtBox(0).Enabled = False
               mVarNoAccount = txtBox(0)
               MyDDE.CancelTrans = False
               MyDDE.GetFieldByName("Default") = 1
               If txtBox(0).Enabled = True Then txtBox(0).SetFocus
            Else
               MessageBox "Data Partner Belum Ada.", "Peringatan"
               MyDDE.CancelTrans = True
            End If
            DGGrid.Enabled = False
       Case tmbAddNew:
            If RcPartner.Recordcount <> 0 Then
               mVarNoAccount = txtBox(0)
               MyDDE.CancelTrans = False
               MyDDE.GetFieldByName("Default") = 1
               If txtBox(0).Enabled = True Then txtBox(0).SetFocus
            Else
               MessageBox "Data Partner Belum Ada.", "Peringatan"
               MyDDE.CancelTrans = True
            End If
            DGGrid.Enabled = False
       Case tmbSave, tmbCancel:
            DGGrid.Enabled = True
End Select
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
PrepareQuery
Err.Clear
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim mDel As New clsDelete
On Error GoTo xErr
Select Case AdReasonActiveDb
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               If mDel.CekDelete(txtBox(0), reDelMasterGudangCust) = False Then
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
       Case tmbEdit:

            If RcPartner.Recordcount <> 0 Then
               MyDDE.CancelTrans = False
            Else
               MessageBox "Data Partner Belum Ada.", "Peringatan"
               MyDDE.CancelTrans = True
            End If
       
       Case tmbAddNew:
            If RcPartner.Recordcount <> 0 Then
               MyDDE.CancelTrans = False
            Else
               MessageBox "Data Partner Belum Ada.", "Peringatan"
               MyDDE.CancelTrans = True
            End If
       Case tmbSave:
            If RcPartner.Recordcount <> 0 Then
                If MyDDE.CheckEmptyControl = False Then
                   'If CekRekening = False Then
                      MyDDE.IsChildMemberReady = True
                      PrepareQuery
                   'Else
                   '   MessageBox "Kode Gudang Sudah ada.", "Peringatan"
                     ' MyDDE.IsChildMemberReady = False
                  ' End If
                Else
                   MyDDE.IsChildMemberReady = False
                End If
            Else
                MessageBox "Data Partner Belum Ada.", "Peringatan"
            End If
End Select
Set mDel = Nothing
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub



Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

Private Sub rcPartner_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
With pRecordset
     If .Recordcount <> 0 Then
        OpenBank IIf(Not IsNull(.Fields("PartnerID")), .Fields("PartnerID"), "XXXX")
     End If
End With
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Function ChekKosong() As Boolean
Dim I As Integer
For I = 0 To 2
    If txtBox(I).Text = "" Then
       ChekKosong = True
       Exit For
    End If
Next I
If cboCurr = "" Then ChekKosong = True
End Function

Private Sub DefCurr()
On Error Resume Next
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
End Sub

Private Sub OpenPartner(ByVal Tipical As String)
CloseDB RcPartner
Set RcPartner = New Recordset
RcPartner.CursorLocation = adUseClient
RcPartner.Open "SELECT     PartnerID, CompanyName FROM PartnerDB WHERE     (PartnerType = N'" & Tipical & "') ORDER BY PartnerID", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
Set RcPartner.ActiveConnection = Nothing
With RcPartner
     Set DGGrid.DataSource = RcPartner
End With
End Sub

Private Sub OpenBank(ByVal NoPartner As String)
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmGudangCust
    .BindFormTAG = "BANK"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT [GDG ID], [Nama Gudang], Alamat, RG, Telp, PartnerID FROM [Gudang Customer]  WHERE (PartnerID = N'" & NoPartner & "') ORDER BY [GDG ID]"
End With
End Sub

Private Sub PrepareQuery()
On Error GoTo xErr
With MyDDE
    .PrepareAppend = " INSERT INTO [Gudang Customer]" & _
                     " ([GDG ID], PartnerID, [Nama Gudang], Alamat, RG, Telp)" & _
                     " VALUES (N'" & ValidString(txtBox(0)) & "', N'" & RcPartner.Fields(0) & "', N'" & ValidString(txtBox(1)) & "', N'" & ValidString(txtBox(2)) & "', N'" & cboCurr.BoundText & "', N'" & ValidString(txtBox(3)) & "')"
                     
    .PrepareUpdate = " UPDATE [Gudang Customer]" & _
                     " Set [Nama Gudang] =N'" & ValidString(txtBox(1)) & "' ,[Alamat] = N'" & ValidString(txtBox(2)) & "', RG = N'" & cboCurr.BoundText & "', Telp = N'" & ValidString(txtBox(3)) & "'" & _
                     " WHERE (IDX = N'" & IdxBank & "') "
                     
    .PrepareDelete = " DELETE FROM [Gudang Customer] WHERE (Idx=N'" & IdxBank & "')"
End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Function IdxBank() As String
Dim RcId As New Recordset
RcId.CursorLocation = adUseClient
RcId.Open "SELECT IDX FROM [Gudang Customer] WHERE     (PartnerID  = N'" & RcPartner.Fields(0) & "') AND ([GDG ID] = N'" & txtBox(0) & "')", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
Set RcId.ActiveConnection = Nothing
IdxBank = "XXX"
With RcId
     If .Recordcount <> 0 Then
        IdxBank = .Fields(0)
     End If
End With
CloseDB RcId
End Function

Private Function CekRekening() As Boolean
Dim Rcb As New DBQuick
Rcb.DBOpen "SELECT PartnerID, [GDG ID],idx FROM [Gudang Customer] WHERE     (PartnerID = N'" & RcPartner.Fields(0) & "') AND ([GDG ID] = N'" & txtBox(0) & "')", CNN, lckLockReadOnly
With Rcb
     If .Recordcount <> 0 Then
        CekRekening = True
     Else
        CekRekening = False
     End If
End With
Rcb.CloseDB
End Function


Private Sub txtBox_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 3 Then
   ValidNum KeyAscii
End If
End Sub

Private Sub GridLayout()
DGGrid.Columns(0).width = 1514.835
DGGrid.Columns(1).width = 1785.26
DataGrid1(0).Columns(0).width = 1514.835
DataGrid1(0).Columns(1).width = 3690.142
DataGrid1(0).Columns(2).width = 2190.047
DataGrid1(0).Columns(3).width = 1514.835
DataGrid1(0).Columns(4).width = 1514.835
End Sub
