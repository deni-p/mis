VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmInvAdj 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adjustment"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9885
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmInvAdj.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   Tag             =   "Inventory Adjustment"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3945
      Left            =   0
      ScaleHeight     =   3945
      ScaleWidth      =   9885
      TabIndex        =   6
      Top             =   0
      Width           =   9885
      Begin VB.OptionButton OptData 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Increase Adjustment"
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
         Height          =   240
         Index           =   0
         Left            =   195
         TabIndex        =   1
         Top             =   180
         Value           =   -1  'True
         Width           =   2400
      End
      Begin VB.OptionButton OptData 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Decrease Adjustment"
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
         Height          =   240
         Index           =   1
         Left            =   2835
         TabIndex        =   2
         Top             =   180
         Width           =   2580
      End
      Begin MSDataGridLib.DataGrid DGPurchase 
         Bindings        =   "FrmInvAdj.frx":6852
         Height          =   2490
         Left            =   120
         TabIndex        =   5
         Top             =   1245
         Width           =   9630
         _ExtentX        =   16986
         _ExtentY        =   4392
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         HeadLines       =   2
         RowHeight       =   15
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "Kode Barang"
            Caption         =   "Kode Barang"
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
            DataField       =   "Nama Barang"
            Caption         =   "Nama Barang"
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
            DataField       =   "Unit Satuan"
            Caption         =   "Satuan"
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
            DataField       =   "QTY Existing"
            Caption         =   "QTY Existing"
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
         BeginProperty Column04 
            DataField       =   "QTY Actual"
            Caption         =   "QTY Actual"
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
         BeginProperty Column05 
            DataField       =   "QTY ADJ"
            Caption         =   "QTY ADJ"
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
         BeginProperty Column06 
            DataField       =   "LokasiGdg"
            Caption         =   "Gudang"
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
         BeginProperty Column07 
            DataField       =   "sl_no"
            Caption         =   "Batch / Lot No"
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
               ColumnWidth     =   2520
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column07 
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "Tanggal Bukti"
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Tag             =   "ASM"
         Top             =   795
         Width           =   2565
         _ExtentX        =   4524
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
         Format          =   16449539
         CurrentDate     =   38272
      End
      Begin VB.Label lblFixAssets 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No Bukti"
         DataField       =   "No Bukti"
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
         Height          =   255
         Index           =   0
         Left            =   1575
         TabIndex        =   3
         Tag             =   "ASM"
         Top             =   510
         Width           =   2565
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   855
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adjustment No."
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
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   7
         Top             =   510
         Width           =   1305
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   150
         X2              =   1600
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   165
         X2              =   1600
         Y1              =   1095
         Y2              =   1095
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmInvAdj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RcDetail As New DBQuick
Private RcPartner As New DBQuick
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private mVarEdit As Boolean
'Private RcDetail As New DBQuick

Private Sub DGPurchase_AfterColEdit(ByVal ColIndex As Integer)
If ColIndex = 5 Then
   If OptData(0).Value = True Then
      MyDDE.ChildRecordset.Fields(4) = MyDDE.ChildRecordset.Fields(3) + MyDDE.ChildRecordset.Fields(5)
   Else
      MyDDE.ChildRecordset.Fields(4) = MyDDE.ChildRecordset.Fields(3) - MyDDE.ChildRecordset.Fields(5)
   End If
Else
End If
End Sub

Private Sub DGPurchase_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If DGPurchase.col >= 5 And mVarEdit = True Then
   DGPurchase.AllowUpdate = True
Else
   DGPurchase.AllowUpdate = False
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE

End Sub

Private Sub Form_Load()
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
OptData(0).BackColor = Picture2.BackColor
OptData(1).BackColor = Picture2.BackColor
'OptData(0).ForeColor = Picture1.BackColor
'OptData(1).ForeColor = Picture1.BackColor
Set mCall = New frmCaller
OpenDB
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
RcDetail.CloseDB
Set RcDetail = Nothing
RcPartner.CloseDB
Set mCall = Nothing
End Sub

Private Sub Form_Resize()
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
'OptData(0).BackColor = Picture2.BackColor
'OptData(1).BackColor = Picture2.BackColor
'OptData(0).ForeColor = Picture1.BackColor
'OptData(1).ForeColor = Picture1.BackColor
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmInvAdj = Nothing
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case TagForm
       Case "Master Barang":
            With MyDDE.ChildRecordset
                 .Fields(0) = mCall.GetFieldByName(0)
                 .Fields(1) = mCall.GetFieldByName(1)
                 .Fields(2) = mCall.GetFieldByName(2)
                 .Fields(3) = IIf(IsEmpty(mCall.GetFieldByName(3)), 0, mCall.GetFieldByName(3))
                 .Fields(4) = IIf(IsEmpty(mCall.GetFieldByName(3)), 0, mCall.GetFieldByName(3))
                 .Fields(5) = 0
                 .Fields(6) = mCall.GetFieldByName("warehouse")
            End With
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim MyData As New clsTransaksi
On Error GoTo xErr
Select Case AdReasonActiveDb
       Case tmbAddNew:
            If OptData(0).Value = True Then
               MyDDE.GetFieldByName("No Bukti") = MyData.PrepareIndex(tmbTransaksiInvADJ, 5, "", TglIndex)
            Else
               MyDDE.GetFieldByName("No Bukti") = MyData.PrepareIndex(tmbTransaksiInvSUB, 5, "", TglIndex)
            End If
            OpenDummy IIf(Not IsNull(MyDDE.GetFieldByName("No Bukti")), MyDDE.GetFieldByName("No Bukti"), "xxxx")
            mVarEdit = True
            DTPicker1.Value = Now
       Case tmbEdit:
            DTPicker1.Enabled = False
            OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("No Bukti")), MyDDE.GetFieldByName("No Bukti"), "xxxx")
            mVarEdit = True
       Case tmbDetail:
            If MyDDE.IsChildMemberReady = True Then
               OpenPartner (0)
               mVarEdit = True
            End If
            
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               With MyDDE.ChildRecordset
                    If .Recordcount <> 0 Then
                       If SendDataToServer("DELETE FROM [Inventory Tabel] WHERE     (RefTrans = N'" & lblFixAssets(0) & "')") = True Then
                            .MoveFirst
                            Do
                              If .EOF = True Then Exit Do
                              If .Fields("QTY ADJ") <> 0 Then
                                 If OptData(0).Value = True Then
                                    SendDataToServer (" INSERT INTO [Inventory Tabel]" & _
                                                      " (NoItem, QTY_IN, RefTrans, DateTrans, TypeTrans,[QTY ADJ],stockTmp,lokasiGdg,sl_no)" & _
                                                      " VALUES     (N'" & .Fields("Kode barang") & "', " & Val(.Fields("QTY Actual")) - Val(.Fields("QTY Existing")) & ", N'" & lblFixAssets(0) & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), N'INVADJ'," & Val(.Fields("QTY Actual")) - Val(.Fields("QTY Existing")) & "," & Val(.Fields("QTY Actual")) - Val(.Fields("QTY Existing")) & ",'" & .Fields("lokasiGdg") & "','" & .Fields("sl_no") & "')")
                                 Else
                                    SendDataToServer (" INSERT INTO [Inventory Tabel]" & _
                                                      "               (NoItem,                        QTY_OUT,                     RefTrans                   , DateTrans                                                          , TypeTrans,[QTY ADJ]      ,stockTmp,lokasiGdg,sl_no)" & _
                                                      " VALUES     (N'" & .Fields("Kode barang") & "', " & .Fields("QTY Adj") & ", N'" & lblFixAssets(0) & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), N'INVSUB',-" & .Fields("QTY Adj") & ",-" & .Fields("QTY Adj") & ",'" & .Fields("lokasiGdg") & "','" & .Fields("sl_no") & "')")
                                 End If
                              End If
                              .MoveNext
                            Loop
                            .MoveLast
                            mVarEdit = False
                       End If
                       OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("No Bukti")), MyDDE.GetFieldByName("No Bukti"), "xxxx")
                    End If
               End With
            End If
       Case tmbPrint:
            CallRPTReport "Inventory ADJ.rpt", "Select * from [Inventory ADJ] Where [No bukti] =N'" & lblFixAssets(0) & "'"
End Select
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("No Bukti")), MyDDE.GetFieldByName("No Bukti"), "xxxx")
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo xErr
Select Case AdReasonActiveDb
       Case tmbAddNew:
            If MyDDE.CheckEmptyControl = False Then
                  MyDDE.IsChildMemberReady = True
                  PrepareQuery
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               'If mDel.CekDelete(txtBox(0), reDelMasterCurency) = False Then
                  MyDDE.IsChildMemberReady = True
                  PrepareQuery
               'Else
               '   MyDDE.CancelTrans = True
               '   MessageBox "Record (" & txtBox(0) & ") Sedang Dipakai Transaksi Lain." & vbCrLf & "Record Tidak Bisa DiHapus.", "Peringatan", msgOkOnly
               '   MyDDE.IsChildMemberReady = False
'               End If
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
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub PrepareQuery()
On Error GoTo xErr
With MyDDE
    .PrepareAppend = " INSERT INTO [Inventory Tabel]" & _
                     " (RefTrans, DateTrans, TypeTrans)" & _
                     " VALUES  (N'" & lblFixAssets(0) & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), N'INVADJ')"

    .PrepareUpdate = " UPDATE    [Inventory Tabel]" & _
                     " SET DateTrans = CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), TypeTrans = N'INVADJ'" & _
                     " WHERE     (RefTrans = N'" & lblFixAssets(0) & "')"

    .PrepareDelete = " DELETE FROM [Inventory Tabel] WHERE   (RefTrans = N'" & lblFixAssets(0) & "') "
End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Function TglIndex() As String
Dim TglHari As String
Dim TglBulan As String
Dim TglTahun As String
If OptData(0).Value = True Then
   TglIndex = "IA-" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
Else
   TglIndex = "IS-" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End If
End Function

Private Sub OpenDetail(ByVal ParamString As String)

If OptData(0).Value = True Then
   RcDetail.DBOpen " SELECT [Inventory Tabel].NoItem AS [Kode Barang], Inventory.InternalNAme AS [Nama Barang], Inventory.UOM AS [Unit Satuan], [Inventory Tabel].[QTY ADJ] AS [QTY Existing], [Inventory Tabel].QTY_IN + [Inventory Tabel].[QTY ADJ] AS [QTY Actual],  [Inventory Tabel].QTY_IN AS [QTY ADJ], [inventory tabel].lokasigdg,[Inventory Tabel].sl_no FROM [Inventory Tabel] INNER JOIN  Inventory ON [Inventory Tabel].NoItem = Inventory.NoItem WHERE     ([Inventory Tabel].RefTrans = N'" & ParamString & "')", CNN
'   RcDetail.DBOpen " SELECT [Inventory Tabel].NoItem AS [Kode Barang], Inventory.InternalName AS [Nama Barang], Inventory.UOM AS [Unit Satuan], [Inventory Tabel].[QTY ADJ] AS [QTY Existing], [Inventory Tabel].[QTY ADJ] AS [QTY Actual]                           ,  [Inventory Tabel].QTY_IN AS [QTY ADJ],  [Inventory Tabel].LokasiGdg,[Inventory Tabel].sl_no FROM [Inventory Tabel] INNER JOIN  Inventory ON [Inventory Tabel].NoItem = Inventory.NoItem WHERE     ([Inventory Tabel].RefTrans = N'" & ParamString & "')", CNN
Else
   RcDetail.DBOpen " SELECT [Inventory Tabel].NoItem AS [Kode Barang], Inventory.InternalName AS [Nama Barang], Inventory.UOM AS [Unit Satuan], [Inventory Tabel].[QTY ADJ] AS [QTY Existing], [Inventory Tabel].[QTY ADJ] AS [QTY Actual], [Inventory Tabel].QTY_IN AS [QTY ADJ], [inventory tabel].lokasigdg,[Inventory Tabel].sl_no FROM  [Inventory Tabel] INNER JOIN                       Inventory ON [Inventory Tabel].NoItem = Inventory.NoItem WHERE     ([Inventory Tabel].RefTrans = N'" & ParamString & "')", CNN
End If

Set MyDDE.ChildRecordset = RcDetail.DBRecordset '.Clone(adLockBatchOptimistic)
Set DGPurchase.DataSource = MyDDE.ChildRecordset
'RcDetail.CloseDB
End Sub

Private Function OpenPartner(ByVal Index As Integer) As Boolean

Select Case Index
       Case 0:
            RcPartner.DBOpen "SELECT [Inventory].NoItem AS [Kode Barang], Inventory.internalName AS [Nama Barang], Inventory.UOM AS [Unit Satuan], SUM([Inventory Tabel].QTY_IN)  - SUM([Inventory Tabel].QTY_OUT) AS Stock, inventory.warehouse FROM [Inventory Tabel] right outer JOIN Inventory ON [Inventory Tabel].NoItem = Inventory.NoItem GROUP BY [Inventory].NoItem, Inventory.InternalName, Inventory.UOM, inventory.warehouse", CNN, lckLockReadOnly
            mCall.FromTagActive = "Master Barang"

End Select
If RcPartner.Recordcount <> 0 Then
    Set mCall.FormData = RcPartner.DBRecordset
    mCall.LookUp Me
    If FindOwnRecordset(MyDDE.ChildRecordset, "[Kode Barang] = '" & mCall.GetFieldByName(0) & "'") = True Then
       MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Kode Barang") & " Sudah Ada....! Silahkan Diulangi", "Peringatan", msgOkOnly, msgCrtical
       'CancelDetailTrans
       'DGPurchase.SetFocus
    End If
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgCrtical
   OpenPartner = True
End If

End Function

Private Sub OpenDB()
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmInvAdj
    '.SetPermissions = UserAddnewDeleteDenied
    .BindFormTAG = "ASM"
    Set .ActiveConnection = CNN
    If OptData(0).Value = True Then
       .PrepareQuery = "SELECT DateTrans AS [Tanggal Bukti], RefTrans AS [No Bukti], TypeTrans FROM         [Inventory Tabel] GROUP BY DateTrans, RefTrans, TypeTrans HAVING      (TypeTrans = N'INVADJ')"
    Else
       .PrepareQuery = "SELECT DateTrans AS [Tanggal Bukti], RefTrans AS [No Bukti], TypeTrans FROM         [Inventory Tabel] GROUP BY DateTrans, RefTrans, TypeTrans HAVING      (TypeTrans = N'INVSUB')"
    End If
End With
End Sub

Private Sub OptData_Click(Index As Integer)
OpenDB
End Sub

Private Sub OpenDummy(ByVal ParamString As String)
'Dim RcDetail As New DBQuick
If OptData(0).Value = True Then
   RcDetail.DBOpen " SELECT [Inventory Tabel].NoItem AS [Kode Barang], Inventory.ItemName AS [Nama Barang], Inventory.UOM AS [Unit Satuan], [Inventory Tabel].[QTY ADJ] AS [QTY Existing], [Inventory Tabel].[QTY ADJ] AS [QTY Actual],  [Inventory Tabel].QTY_IN AS [QTY ADJ],  [Inventory Tabel].LokasiGdg,[Inventory Tabel].sl_no FROM [Inventory Tabel] INNER JOIN  Inventory ON [Inventory Tabel].NoItem = Inventory.NoItem WHERE     ([Inventory Tabel].RefTrans = N'" & "ParamString" & "')", CNN
Else
   RcDetail.DBOpen " SELECT [Inventory Tabel].NoItem AS [Kode Barang], Inventory.ItemName AS [Nama Barang], Inventory.UOM AS [Unit Satuan], [Inventory Tabel].[QTY ADJ] AS [QTY Existing], [Inventory Tabel].[QTY ADJ] AS [QTY Actual], [Inventory Tabel].QTY_OUT AS [QTY ADJ],  [Inventory Tabel].LokasiGdg,[Inventory Tabel].sl_no  FROM         [Inventory Tabel] INNER JOIN                       Inventory ON [Inventory Tabel].NoItem = Inventory.NoItem WHERE     ([Inventory Tabel].RefTrans = N'" & "ParamString" & "')", CNN
End If
Set MyDDE.ChildRecordset = RcDetail.DBRecordset '.Clone(adLockBatchOptimistic)
Set DGPurchase.DataSource = MyDDE.ChildRecordset
'RcDetail.CloseDB
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

