VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00F8C996-2DE8-46A8-BC86-FC76BF56E773}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmMPS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MPS"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11670
   Icon            =   "frmMPS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   11670
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   0
      ScaleHeight     =   6855
      ScaleWidth      =   11670
      TabIndex        =   11
      Top             =   0
      Width           =   11670
      Begin VB.ComboBox Combo1 
         DataField       =   "issue_to"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMPS.frx":6852
         Left            =   1485
         List            =   "frmMPS.frx":6859
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Tag             =   "PO"
         Top             =   1950
         Width           =   1905
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "description"
         Height          =   330
         Index           =   1
         Left            =   1485
         MaxLength       =   15
         TabIndex        =   2
         Tag             =   "PO"
         Top             =   480
         Width           =   3450
      End
      Begin VB.ComboBox CmbPeriod 
         Appearance      =   0  'Flat
         DataField       =   "periode_type"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMPS.frx":6867
         Left            =   3030
         List            =   "frmMPS.frx":6874
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "PO"
         Top             =   840
         Width           =   1905
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "no_mps"
         Height          =   330
         Index           =   0
         Left            =   1485
         MaxLength       =   15
         TabIndex        =   1
         Tag             =   "PO"
         Top             =   135
         Width           =   3450
      End
      Begin VB.TextBox txtIsi 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   2520
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1005
         ItemData        =   "frmMPS.frx":688C
         Left            =   9720
         List            =   "frmMPS.frx":688E
         TabIndex        =   10
         Top             =   2520
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtperiode 
         Appearance      =   0  'Flat
         DataField       =   "periode_no"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   1490
         MaxLength       =   15
         TabIndex        =   3
         Tag             =   "PO"
         Top             =   840
         Width           =   1530
      End
      Begin MSComCtl2.DTPicker FCastDate 
         DataField       =   "End Date"
         Height          =   330
         Index           =   1
         Left            =   1485
         TabIndex        =   6
         Tag             =   "PO"
         Top             =   1581
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   582
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
         Format          =   64290819
         CurrentDate     =   38649
      End
      Begin MSComCtl2.DTPicker FCastDate 
         DataField       =   "Require Date"
         Height          =   330
         Index           =   0
         Left            =   1485
         TabIndex        =   5
         Tag             =   "PO"
         Top             =   1200
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   582
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
         Format          =   64290819
         CurrentDate     =   38649
      End
      Begin MSDataGridLib.DataGrid DGPurchase 
         Height          =   4335
         Left            =   120
         TabIndex        =   8
         Tag             =   "Partner"
         Top             =   2400
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   7646
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16777215
         HeadLines       =   2
         RowHeight       =   15
         RowDividerStyle =   6
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
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   135
         X2              =   1635
         Y1              =   2250
         Y2              =   2250
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue To"
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
         Index           =   5
         Left            =   135
         TabIndex        =   17
         Top             =   2010
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Index           =   9
         Left            =   135
         TabIndex        =   16
         Top             =   555
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period"
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
         Index           =   4
         Left            =   135
         TabIndex        =   15
         Top             =   915
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
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
         Left            =   135
         TabIndex        =   14
         Top             =   1650
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MPS Name"
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
         Left            =   135
         TabIndex        =   13
         Top             =   195
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
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
         Left            =   135
         TabIndex        =   12
         Top             =   1290
         Width           =   855
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   135
         X2              =   1560
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   135
         X2              =   1560
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   135
         X2              =   1560
         Y1              =   1155
         Y2              =   1155
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   135
         X2              =   1560
         Y1              =   1530
         Y2              =   1530
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   135
         X2              =   1560
         Y1              =   1890
         Y2              =   1890
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   6930
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmMPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RcJadwal As New Recordset
Private clsMytr As New DBQuick
Private RcUang As New DBQuick
Private RcDetail As New DBQuick
Private RcDetailTest As New Recordset
Private RcPartner As New DBQuick
Private mAwal As Integer
Private mAkhir As Integer
Private mCount As Integer
Private mList As String
Private mRowLast As Long
Private Rctest As New Recordset
Private RcTestIsi As New Recordset
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private MyData As New clsTransaksi
Dim RsDetail As DBQuick
Dim ChekBaris, ChekBarisDel As Integer
Private MEdit, mEditPO, mFirstCaller As Boolean






Private Sub DGPurchase_Click()
txtIsi.Visible = False
End Sub

Private Sub DGPurchase_DblClick()
On Error Resume Next
If DGPurchase.Col >= 2 Then
    txtIsi.Move DGPurchase.Columns(DGPurchase.Col).Left + 120, DGPurchase.RowTop(DGPurchase.Row) + 2400, DGPurchase.Columns(DGPurchase.Col).Width, DGPurchase.RowHeight
    txtIsi.Visible = True
    txtIsi.Text = MyDDE.ChildRecordset(DGPurchase.Col)
    txtIsi.SetFocus
End If
End Sub

Private Sub FCastDate_Change(Index As Integer)
If Index = 1 Then
   mList = CmbPeriod.Text
   txtperiode(0).Text = HitungHari
   GenerateJadwal 0, txtperiode(0)
End If
End Sub

Private Sub FCastDate_Click(Index As Integer)
If Index = 1 Then
   mList = CmbPeriod.Text
   txtperiode(0).Text = HitungHari
   GenerateJadwal 0, txtperiode(0)
End If

End Sub

Private Sub Form_Activate()
'digunakan membaca awal banyaknya baris di grid
If MyDDE.ChildRecordset.Recordcount > 0 Then
    MyDDE.ChildRecordset.MoveLast
    ChekBaris = DGPurchase.Row
End If
End Sub

Private Sub Form_Load()
'GridLayout
HiasFormManTell Picture2, Me
Set mCall = New frmCaller
FCastDate(0).Value = dDateBegin
FCastDate(1).Value = dDateBegin
With MyDDE
     .EditModeReplace = False
      Set .BindForm = frmMPS
      Set .ActiveConnection = CNN
     .BindFormTAG = "PO"
     .PrepareQuery = "SELECT * from [mps header]"
End With
GenerateJadwal 0, 0
Load_FCast_Item
list_dataGrid MyDDE.GetFieldByName("no_mps") 'digunakan untuk tampilkan data di grid

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set MyData = Nothing
MyDDE.ClearRecordset
'RcUang.CloseDB
'clsMytr.CloseDB
'RcPartner.CloseDB
Set mCall = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmSalesForecast = Nothing
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Dim i As Integer
Select Case TagForm:
       Case "MASTER BARANG":
             For i = 0 To 2   'digunakan input list 3
                MyDDE.ChildRecordset("Item No") = mCall.GetFieldByName(0)
                MyDDE.ChildRecordset.AddNew
             Next i
             MyDDE.ChildRecordset.Delete
             insert_Item_Fcast 'digunakan insert Sales_Fcast_Item yang di ambil dari tabel
             MyDDE.ChildRecordset.MoveLast
             ChekBaris = DGPurchase.Row
End Select
End Sub

Private Sub MYDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
'txtBox(0).Enabled = False
txtperiode(0).Enabled = False
Select Case AdReasonActiveDb
       Case tmbEdit:
             MEdit = True
             MyDDE.ChildRecordset.MoveLast
       Case tmbAddNew:
             MEdit = False
             clear_grid  'di gunakan clear grid untuk tambah data
       Case tmbSave:
          If MyDDE.IsChildMemberReady = True Then
               SimpanDetail True
               MEdit = False
               mEditPO = False
            End If
       Case tmbCancel:
            If MyDDE.ChildRecordset.Recordcount = 0 Then
               list_dataGrid MyDDE.GetFieldByName("no_mps")
               MEdit = False
            Else
               MEdit = True
            End If
       Case tmbDetail:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               OpenPartner 1
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbNextRecord:
            list_dataGrid txtBox(0)
            Form_Activate
       Case tmbPreviousRecord:
            list_dataGrid txtBox(0)
            Form_Activate
       Case tmbBottomRecord:
            list_dataGrid txtBox(0)
            Form_Activate
       Case tmbTopRecord:
            list_dataGrid txtBox(0)
            Form_Activate
       Case tmbPrint:
          '  CallRPTReport "Sales Contract.rpt", "Select * From [Sales Contract] where PurchaseID ='" & txtBox(0) & "'"
       Case tmbQuit:
           ' Unload Me

End Select
GridLayout
Err.Clear
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub PrepareQuery()
With MyDDE
    .PrepareAppend = " INSERT INTO  [MPS Header] ( No_MPS, description, periode_no, Periode_type, [Require Date], [end date], issue_to) " & _
                     " VALUES (N'" & txtBox(0) & "', N'" & txtBox(1) & "'," & txtperiode(0).Text & ",'" & CmbPeriod.Text & "',convert(Datetime, '" & Format(FCastDate(0).Value, "dd/mm/yy") & "',3),convert(Datetime,'" & Format(FCastDate(1).Value, "dd/mm/yy") & "',3),'" & Combo1 & "')"

    .PrepareUpdate = "Update [MPS Header]" & _
                     " set description='" & txtBox(1) & "', periode_no=" & txtperiode(0).Text & ",periode_type='" & CmbPeriod.Text & "' , [Require Date]=convert(Datetime, '" & Format(FCastDate(0).Value, "dd/mm/yy") & "',3), [end date]=convert(Datetime,'" & Format(FCastDate(1).Value, "dd/mm/yy") & "',3),issue_to='" & Combo1.Text & "'" & _
                     " where no_mps='" & MyDDE.GetFieldByName("no_mps") & "'"
    .PrepareDelete = "delete from [MPS Header] where no_mps = '" & MyDDE.GetFieldByName("no_mps") & "'"
End With
End Sub


Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:
Set RcPartner = New DBQuick
Select Case Index
       Case 1:
             RcPartner.DBOpen "SELECT NoItem AS [BOM Id], ItemName AS Keterangan, UOM AS UOM FROM Inventory WHERE     (Manufacture = 1) ORDER BY NoItem", CNN, lckLockReadOnly
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 1:
            mCall.FromTagActive = "MASTER BARANG"
            mCall.CaptionLink = "Barang"
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me
End If
Exit Sub
Hell:
    Err.Clear
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
     Case tmbEdit:
     Case tmbDelete:
          MyDDE.ChildRecordset.CancelBatch
     Case tmbDetail: '
     Case tmbSave:
        If MyDDE.CheckEmptyControl = False Then
           If MyDDE.ChildRecordset.Recordcount <> 0 Then
              MyDDE.IsChildMemberReady = True
           Else
              MyDDE.IsChildMemberReady = False
           End If
        Else
           MyDDE.IsChildMemberReady = False
        End If
    Case tmbCancel:
        MyDDE.ChildRecordset.CancelBatch
End Select
End Sub

Private Sub SimpanDetail(ByVal Tipical As Boolean)
On Error Resume Next
Dim i, j As Integer

If MEdit = True Then
With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
           i = 1 ' colom
           Do While i <= FQty(txtperiode(0).Text) + 1
              DGPurchase.Col = i
              j = 0
              Do While j <= ChekBaris 'DGPurchase.Row '2
                DGPurchase.Col = i
                DGPurchase.Row = j
                SendDataToServer "DELETE FROM [MPS Detail] WHERE     (no_mps = N'" & txtBox(0) & "')"

                j = j + 1
              Loop
            i = i + 1
          Loop
        '  End If
    End If
End With
End If


With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
           i = 1 ' colom
           Do While i <= FQty(txtperiode(0).Text) + 1
              DGPurchase.Col = i
              j = 0
              Do While j <= ChekBaris 'DGPurchase.Row '2
                DGPurchase.Col = i
                DGPurchase.Row = j
                SendDataToServer " INSERT INTO [mps detail] ( no_mps, noitem, fcast_item,no_urut,list_value1)" & _
                                 " VALUES (N'" & txtBox(0) & "','" & DGPurchase.Columns(0).Value & "' ,'" & DGPurchase.Columns(1).Value & "'," & FQty(DGPurchase.Columns(1 + i).Caption) & " ," & FQty(DGPurchase.Columns(1 + i).Value) & "  )"

                j = j + 1
              Loop
            i = i + 1
          Loop
    End If
End With
End Sub

Private Sub GridLayout()
Dim i As Integer
If DGPurchase.Splits.Count < 2 Then DGPurchase.Splits.Add 1
DGPurchase.Splits(0).Columns(0).Visible = True
DGPurchase.Splits(0).Columns(1).Visible = True
DGPurchase.Splits(0).Columns(1).Width = 1750
DGPurchase.Splits(1).Columns(0).Visible = False
DGPurchase.Splits(1).Columns(1).Visible = False
DGPurchase.Splits(0).ScrollBars = dbgHorizontal
DGPurchase.Splits(1).ScrollBars = dbgBoth
'DGPurchase.Splits(1).RecordSelectors = False
DGPurchase.Splits(0).AllowSizing = False
DGPurchase.Splits(1).Size = 2
DGPurchase.Splits(1).SizeMode = dbgScalable
'DGPurchase.Splits(1).LeftCol = 0
'DGPurchase.Splits(0).Columns(0).Caption = "No Item"
For i = 0 To RcJadwal.Fields.Count - 1
    If i <= 1 Then
       DGPurchase.Columns(i).Width = 1700
       DGPurchase.Columns(i).Alignment = dbgLeft
    Else
       DGPurchase.Columns(i).Width = 600
       DGPurchase.Columns(i).Alignment = dbgRight
       DGPurchase.Columns(i).NumberFormat = "#,##0;(#,##0)"
       DGPurchase.Splits(0).Columns(i).Visible = False
       DGPurchase.Splits(1).Columns(i).AllowSizing = True
    End If
    DGPurchase.Columns(i).DividerStyle = dbgRaised
Next
DGPurchase.TabAcrossSplits = True
DGPurchase.Refresh


End Sub


Private Sub GenerateJadwal(ByVal vAwal As Integer, ByVal vAkhir As Integer)
On Error Resume Next
Dim i As Integer
Set RcJadwal = Nothing
Set RcJadwal = New Recordset
With RcJadwal
     .Fields.Append "Item No", adBSTR
     .Fields.Append "FCast Item", adBSTR
     '.Fields.Append "Stock", adInteger
     For i = 1 To vAkhir
        .Fields.Append vAwal + i, adInteger
     Next
End With
RcJadwal.Open
Set DGPurchase.DataSource = RcJadwal
mAwal = vAwal
mAkhir = RcJadwal.Fields.Count
mCount = vAkhir
GridLayout
Set MyDDE.ChildRecordset = RcJadwal
Err.Clear
End Sub

Private Sub OpenCucakRowo(ByVal vKode As String, Optional ByVal Tipical As Boolean = False)
On Error Resume Next
Dim Rc As New DBQuick
Dim RcDetail As New Recordset
Dim vWeek As Integer
Dim i As Integer
Dim iJ As Integer
Dim mLast As Integer
Dim Avdata As Variant
'Dim mStart As Boolean
mRowLast = 0
'mTotal = CDbl(txtMrp(6))
vWeek = Format(CDate(IIf(Not IsNull(MyDDE.GetFieldByName("Require Date")), MyDDE.GetFieldByName("Require Date"), Date)), "ww")
mList = IIf(Not IsNull(MyDDE.GetFieldByName("Plan Horizon")), MyDDE.GetFieldByName("Plan Horizon"), "Week")
Select Case MyDDE.GetFieldByName("Plan Horizon")
       Case "Day": GenerateJadwal (Day(MyDDE.GetFieldByName("Require Date"))), HitungHari
       Case "Week": GenerateJadwal vWeek, HitungHari
'       Case "Monthly": GenerateJadwal MonthOfYear(DTPicker1(0)), HitungHari
       Case Else: GenerateJadwal vWeek, HitungHari
End Select
'Tolong Jarno ben sak garis wae, debug-e dhek sql cek enak....Please.
If Tipical = False Then
  ' Rc.DBOpen " SELECT [BOM Component Detail].Component, Inventory.ItemName, Inventory.LeadTimeDays AS LeadTimeDays,  ISNULL([Inventory Tabel].QTY_IN - [Inventory Tabel].QTY_OUT, 0) AS QTY, [BOM Component Detail].QTYUsage FROM  [BOM Component Detail] INNER JOIN  Inventory ON [BOM Component Detail].Component = Inventory.NoItem LEFT OUTER JOIN                       [Inventory Tabel] ON Inventory.NoItem = [Inventory Tabel].NoItem WHERE  ([BOM Component Detail].NoItem = N'" & vKode & "') ORDER BY [BOM Component Detail].Component", CNN, lckLockReadOnly
   With Rc.DBRecordset
        If .Recordcount <> 0 Then
        'messagebox .Source
           'LoadJadwal MyDDE.GetFieldByName("NoItem"), MyDDE.GetFieldByName("Lead Time"), MyDDE.GetFieldByName("QTY") - MyDDE.GetFieldByName("Unit Aloc"), IIf(Not IsNull(MyDDE.GetFieldByName("Safety Stock")), MyDDE.GetFieldByName("Safety Stock"), 0), True
           Do
              If .EOF = True Then Exit Do
                ' LoadJadwal .Fields(0), .Fields(2), .Fields(3), IIf(Not IsNull(MyDDE.GetFieldByName("Safety Stock")), MyDDE.GetFieldByName("Safety Stock"), 0), True '.Fields(3)
                 mRowLast = mRowLast + 6
                .MoveNext
           Loop
        End If
   End With
Else
 ' Rc.DBOpen "Shape{SELECT [MRP Detail].Component, [MRP INVENTORY].[Require Date], [MRP INVENTORY].[End Date] FROM [MRP INVENTORY] INNER JOIN [MRP Detail] ON [MRP INVENTORY].NoItem = [MRP Detail].NoItem WHERE ([MRP INVENTORY].NoItem = N'" & vKode & "') GROUP BY [MRP INVENTORY].[Require Date], [MRP INVENTORY].[End Date], [MRP Detail].Component ORDER BY [MRP Detail].Component} Append({SELECT [MRP Detail].Component AS Component, [MRP Detail].[Time Days] AS [Plan Horizon], [MRP Detail].[List Value1] AS Amount, [MRP Detail].[No Urut] FROM [MRP Detail] INNER JOIN Inventory ON [MRP Detail].NoItem = Inventory.NoItem WHERE     (Inventory.Manufacture = 1) AND ([MRP Detail].NoItem = N'" & vKode & "') GROUP BY [MRP Detail].[Time Days], [MRP Detail].[List Value1], [MRP Detail].[No Urut], [MRP Detail].Component ORDER BY [MRP Detail].Component, [MRP Detail].[No Urut]} As ChildMD Relate Component To Component)", CNN, lckLockBatch
  With Rc.DBRecordset
       If .Recordcount <> 0 Then
'          vWeek = Format(CDate(IIf(Not IsNull(MyDDE.GetFieldByName("Require Date")), MyDDE.GetFieldByName("Require Date"), Date)), "ww")
'          mList = MyDDE.GetFieldByName("Plan Horizon")
'          Select Case MyDDE.GetFieldByName("Plan Horizon")
'                 Case "Day": GenerateJadwal (Day(MyDDE.GetFieldByName("Require Date"))), HitungHari
'                 Case "Week": GenerateJadwal vWeek, HitungHari
''                 Case "Monthly": GenerateJadwal MonthOfYear(DTPicker1(0)), HitungHari
'                 Case Else: GenerateJadwal vWeek, HitungHari
'          End Select
          Set RcDetail = Rc.DBRecordset("ChildMD").UnderlyingValue
'            RcDetail.MoveFirst
'            messagebox RcDetail.GetString(adClipString)
'             LoadJadwal MyDDE.GetFieldByName("NoItem"), MyDDE.GetFieldByName("Lead Time"), MyDDE.GetFieldByName("QTY") - MyDDE.GetFieldByName("Unit Alloc"), IIf(Not IsNull(MyDDE.GetFieldByName("Safety Stock")), MyDDE.GetFieldByName("Safety Stock"), 0), True
             Do
               If .EOF Then Exit Do
                  mLast = 1
                  iJ = 0
                  If RcDetail.Recordcount <> 0 Then
                        Avdata = RcDetail.Getrows(RcDetail.Recordcount, adBookmarkFirst)
                        For i = 0 To UBound(Avdata, 2)
                            iJ = iJ + 1
                            If mLast <> Avdata(3, i) Then iJ = 1
                            Select Case Avdata(3, i)
                                   Case 1:
                                        If iJ = 1 Then RcJadwal.AddNew: RcJadwal.Fields(0) = .Fields(0): RcJadwal.Fields(1) = "Gross Requirement"
                                   Case 2:
                                        If iJ = 1 Then RcJadwal.AddNew: RcJadwal.Fields(1) = "Schedule Receipt"
                                   Case 3:
                                        If iJ = 1 Then RcJadwal.AddNew: RcJadwal.Fields(1) = "On Hand"
                                   Case 4:
                                        If iJ = 1 Then RcJadwal.AddNew: RcJadwal.Fields(1) = "Net Requirement"
                                   Case 5:
                                        If iJ = 1 Then RcJadwal.AddNew: RcJadwal.Fields(1) = "Plan Order Receipt"
                                   Case 6:
                                        If iJ = 1 Then RcJadwal.AddNew: RcJadwal.Fields(1) = "Plan Order Release"
                            End Select
'                            messagebox RcJadwal.Fields(10).Name
                            RcJadwal.Fields((2 + iJ) - 1) = Avdata(2, i)
                            'RcJadwal.Fields(iJ - 1) = Avdata(2, I)
                            mLast = Avdata(3, i)
                        Next i
                  End If
               .MoveNext
             Loop
             .MoveFirst
       Else
           
       End If
  End With
End If
End Sub


Private Function HitungHari() As Long
Dim mTotal As Long
'mTotal = CDate(IIf(Not IsNull(MyDDE.GetFieldByName("Require Date")), MyDDE.GetFieldByName("Require Date"), Date)) - CDate(IIf(Not IsNull(MyDDE.GetFieldByName("End Date")), MyDDE.GetFieldByName("End Date"), Date))
'mTotal = CDate(IIf(Not IsNull(FCastDate(0).value), FCastDate(0).value, Date)) - CDate(IIf(Not IsNull(FCastDate(1).value), FCastDate(1).value, Date))
mTotal = CDate(IIf(Not IsNull(FCastDate(1).Value), FCastDate(1).Value, Date)) - CDate(IIf(Not IsNull(FCastDate(0).Value), FCastDate(0).Value, Date))
Select Case mList
       Case "DAY": HitungHari = mTotal
       Case "WEEK": HitungHari = Round(mTotal / 7)
       Case "MONTHLY": HitungHari = Round(mTotal / 30)
End Select
End Function

Private Sub insert_FCast_Item()
Dim i As Integer
Set RcDetailTest = New ADODB.Recordset
RcDetailTest.Open "select * from sales_fcast_item", CNN
RcDetailTest.MoveFirst
If RcDetailTest.Recordcount > 0 Then
    i = 0
    Do While RcDetailTest.EOF = False
        DGPurchase.Row = i
        MyDDE.ChildRecordset(1) = RcDetailTest.Fields("fcast_item")
        RcDetailTest.MoveNext
        i = i + 1
    Loop
End If
End Sub

Private Sub Load_FCast_Item()
Dim i As Integer
Set RcDetailTest = New ADODB.Recordset
RcDetailTest.Open "select * from sales_fcast_item", CNN
RcDetailTest.MoveFirst
If RcDetailTest.Recordcount > 0 Then
    Do While RcDetailTest.EOF = False
        List1.AddItem RcDetailTest.Fields("fcast_item")
        RcDetailTest.MoveNext
    Loop
End If
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   MyDDE.ChildRecordset(DGPurchase.Col) = txtIsi.Text
   txtIsi.Visible = False
   txtIsi.Text = ""
End If
End Sub

Private Sub insert_Item_Fcast()
Dim i, j As Integer
 j = 0
 For i = 0 To DGPurchase.Row
    DGPurchase.Row = i
    MyDDE.ChildRecordset(1) = List1.List(j)
    j = j + 1
    If j = 3 Then
       j = 0
    End If
 Next i
End Sub

Private Sub list_dataGrid(idfcast As String)
Dim i, j, h As Integer

Set Rctest = New Recordset
Set RcTestIsi = New Recordset

If txtperiode(0).Text <> "" Then
    GenerateJadwal 0, FQty(txtperiode(0).Text)
    
'    Rctest.Open "SELECT  dbo.sales_fcast_line.item_no, dbo.sales_fcast_line.fcast_item " & _
'                "FROM  dbo.sales_fcast INNER JOIN " & _
'                "dbo.sales_fcast_line ON dbo.sales_fcast.fcast_id = dbo.sales_fcast_line.fcast_id INNER JOIN " & _
'                "dbo.sales_fcast_item ON dbo.sales_fcast_line.fcast_item = dbo.sales_fcast_item.fcast_item " & _
'                " where dbo.sales_fcast_line.fcast_id='" & idfcast & "'" & _
'                "GROUP BY dbo.sales_fcast_line.item_no, dbo.sales_fcast_line.fcast_item", CNN, adOpenKeyset, adLockReadOnly
                
    
    Rctest.Open "SELECT dbo.[MPS Detail].NoItem, dbo.[MPS Detail].fcast_item " & _
                "FROM dbo.[MPS Header] INNER JOIN " & _
                "dbo.[MPS Detail] ON dbo.[MPS Header].No_MPS = dbo.[MPS Detail].No_MPS INNER JOIN " & _
                "dbo.sales_fcast_item ON dbo.[MPS Detail].fcast_item = dbo.sales_fcast_item.fcast_item " & _
                "Where dbo.[MPS Detail].No_MPS ='" & idfcast & "'" & _
                "GROUP BY dbo.[MPS Detail].NoItem, dbo.[MPS Detail].fcast_item", CNN, adOpenKeyset, adLockReadOnly
    
    
    
'    RcTestIsi.Open "SELECT  dbo.sales_fcast_line.fcast_item, dbo.sales_fcast_line.list_value1, dbo.sales_fcast_line.item_no, dbo.sales_fcast_line.no_urut " & _
'                   "FROM    dbo.sales_fcast INNER JOIN " & _
'                   "dbo.sales_fcast_line ON dbo.sales_fcast.fcast_id = dbo.sales_fcast_line.fcast_id INNER JOIN " & _
'                   "dbo.sales_fcast_item ON dbo.sales_fcast_line.fcast_item = dbo.sales_fcast_item.fcast_item where dbo.sales_fcast_line.fcast_id = '" & idfcast & "'" & _
'                   "GROUP BY dbo.sales_fcast_line.fcast_item, dbo.sales_fcast_line.list_value1, dbo.sales_fcast_line.item_no, dbo.sales_fcast_line.no_urut", CNN, adOpenKeyset, adLockReadOnly


     RcTestIsi.Open "SELECT dbo.[MPS Detail].fcast_item, dbo.[mps detail].list_value1,dbo.[MPS Detail].NoItem,dbo.[MPS Detail].no_urut " & _
                "FROM dbo.[MPS Header] INNER JOIN " & _
                "dbo.[MPS Detail] ON dbo.[MPS Header].No_MPS = dbo.[MPS Detail].No_MPS INNER JOIN " & _
                "dbo.sales_fcast_item ON dbo.[MPS Detail].fcast_item = dbo.sales_fcast_item.fcast_item " & _
                "Where dbo.[MPS Detail].No_MPS ='" & idfcast & "'" & _
                "GROUP BY dbo.[MPS Detail].fcast_item, dbo.[mps detail].list_value1, dbo.[MPS Detail].NoItem, dbo.[MPS Detail].no_urut", CNN, adOpenKeyset, adLockReadOnly


    h = 0
    Do While h <= Rctest.Recordcount - 1   'baris
        MyDDE.ChildRecordset.AddNew
    h = h + 1
    Loop
    
    With MyDDE.ChildRecordset
         If .Recordcount <> 0 Then
               i = 1 ' colom
               Do While i <= FQty(txtperiode(0).Text) + 1 'kolom
                  DGPurchase.Col = i
                  j = 0
                  Rctest.MoveFirst
                  Do While j <= Rctest.Recordcount - 1 'DGPurchase.Row '2 baris
                     DGPurchase.Col = i
                     DGPurchase.Row = j '
                     MyDDE.ChildRecordset("item no") = Rctest.Fields("noitem")
                     MyDDE.ChildRecordset("fcast item") = Rctest.Fields("fcast_item")
                     j = j + 1
                     Rctest.MoveNext
                  Loop
                i = i + 1
              Loop
        End If
    End With
    
    
    'digunakan isi atau nampilkan data di grid
    With MyDDE.ChildRecordset
         If RcTestIsi.Recordcount <> 0 Then
            i = 0
            Do While i <= Rctest.Recordcount - 1 'FQty(txtperiode(0).text) + 1
              ' DGPurchase.Col = I
               DGPurchase.Row = i
               j = 1
               Do While j <= FQty(txtperiode(0).Text) + 1
                  RcTestIsi.MoveFirst
                  Do While RcTestIsi.EOF = False
                     DGPurchase.Row = i '
                     DGPurchase.Col = j
                     If DGPurchase.Columns(0).Value = RcTestIsi.Fields("noitem") And Trim(DGPurchase.Columns(1).Value) = Trim(RcTestIsi.Fields("Fcast_item")) And DGPurchase.Columns(j).Caption = RcTestIsi.Fields("no_urut") Then
                        MyDDE.ChildRecordset(j) = IIf(IsNull(RcTestIsi.Fields("list_value1") = True), 0, RcTestIsi.Fields("list_value1"))
                     End If
                     RcTestIsi.MoveNext
                  Loop
                  j = j + 1
                Loop
                i = i + 1
                MyDDE.ChildRecordset.MoveNext
            Loop
        End If
    End With
End If

Set Rctest = Nothing
Set RcTestIsi = Nothing
End Sub

Private Sub clear_grid()
Dim i As Integer
MyDDE.ChildRecordset.MoveLast
i = DGPurchase.Row
Do While i >= 0 'baris
    DGPurchase.Row = i
    MyDDE.ChildRecordset.Delete
    i = i - 1
Loop

End Sub



