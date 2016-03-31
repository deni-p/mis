VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmKirimProduk 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Serah terima Produk Jadi"
   ClientHeight    =   6105
   ClientLeft      =   1635
   ClientTop       =   1920
   ClientWidth     =   11190
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
   Icon            =   "FrmSetorProdukJadi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   Tag             =   "Purchase Order"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5535
      ScaleWidth      =   11190
      TabIndex        =   3
      Top             =   0
      Width           =   11190
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "datetrans"
         DataSource      =   "mydde"
         Height          =   315
         Left            =   1590
         TabIndex        =   6
         Top             =   495
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   71630851
         CurrentDate     =   39618
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "IDTrans"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   0
         Left            =   1590
         MaxLength       =   16
         TabIndex        =   1
         Tag             =   "FG"
         Top             =   150
         Width           =   2400
      End
      Begin MSDataGridLib.DataGrid DGPurchase 
         Height          =   4410
         Left            =   75
         TabIndex        =   2
         Top             =   990
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   7779
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   1
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "noItem"
            Caption         =   "No Item"
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
            DataField       =   "itemName"
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
            DataField       =   "ln_No"
            Caption         =   "Lot No"
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
            DataField       =   "kuantitas"
            Caption         =   "Kuantitas"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "satuan"
            Caption         =   "Satuan"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "keterangan"
            Caption         =   "Keterangan"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "#,##0.00;(#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            Size            =   248
            BeginProperty Column00 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column05 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Transaksi"
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
         Left            =   120
         TabIndex        =   5
         Top             =   210
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
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
         Left            =   120
         TabIndex        =   4
         Top             =   555
         Width           =   645
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   150
         X2              =   1650
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   150
         X2              =   1650
         Y1              =   795
         Y2              =   795
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5535
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   1005
      BindFormTAG     =   "FG"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmKirimProduk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RcDetail                                          As New DBQuick
Attribute RcDetail.VB_VarHelpID = -1
Private RcLookup                                          As New DBQuick
Private WithEvents mCall                                  As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private IDGen As New IDGenerator


Private Sub cmdLink_Click(Index As Integer)
   OpenPartner Index
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
   HiasFormManTell Picture2, Me
   Set mCall = New frmCaller
   'DTPicker1.Value = dDateBegin
   With MyDDE
        .EditModeReplace = False
        .BindFormTAG = "FG"
        Set .BindForm = Me
        Set .ActiveConnection = CNN
'        .PrepareQuery = "SELECT * from finishGood where received_by = ''"
        .PrepareQuery = "SELECT * from backflush_header where status = 0 and TypeTrans = 'FG'"
        .SetPermissions = aksess.MayDo("Penerimaan Barang Jadi")
   End With
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If pRecordset.Recordcount <> 0 Then
   Select Case TagForm:
      Case "Produk Dikirim":
         MyDDE.ChildRecordset.Fields("NoItem") = mCall.GetFieldByName("NoItem")
         MyDDE.ChildRecordset.Fields("ItemName") = mCall.GetFieldByName("ItemName")
         MyDDE.ChildRecordset.Fields("ln_no") = mCall.GetFieldByName("lotno")
         MyDDE.ChildRecordset.Fields("kuantitas") = mCall.GetFieldByName("hasil_powder")
         MyDDE.ChildRecordset.Fields("satuan") = mCall.GetFieldByName("UOM")
         MyDDE.ChildRecordset.Fields("keterangan") = "-"
   End Select
End If
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
   PrepareQuery
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit:
       Case tmbDetail:
       Case tmbSave:
         If MyDDE.ChildRecordset.Recordcount > 0 Then
            MyDDE.IsChildMemberReady = True
         Else
            MyDDE.IsChildMemberReady = False
         End If
      Case tmbDelete:
            With MyDDE.ChildRecordset
               If .Recordcount > 0 Then
                  .MoveFirst
                  While Not .EOF
                     SendDataToServer "Update blending_header set status=0 where lotNo ='" & .Fields("ln_no") & "'"
                     .MoveNext
                  Wend
               End If
            End With
            
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim IDGen As New IDGenerator
Dim rsBalance As New DBQuick
   Select Case AdReasonActiveDb
          Case tmbAddNew:
               DTPicker1.Value = Now
               MyDDE.GetFieldByName("IDTrans") = IDGen.GetID("FG")
               MyDDE.GetFieldByName("Tanggal") = Now
               DTPicker1.SetFocus
          
          Case tmbSave:
               If MyDDE.IsChildMemberReady = True Then
                  SimpanDetail
               Else
                  MessageBox "Detail Item  belum ada datanya.", "Peringatan", msgOkOnly, msgCrtical
               End If
               
         Case tmbDetail:
                  OpenPartner 0
                 
         Case tmbPrint:
            Dim lPrint As New utility
            lPrint.CallReportView "select * from serah_terima_produk_jadi where IDTrans='" & txtBox(0).Text & "'", "SERAH TERIMA PRODUK JADI.rpt", ReportPath, "Serah Terima Produk Jadi"
            Set lPrint = Nothing
         
   End Select
End Sub

Private Sub UpdateBlending(strLotNo As String)
   SendDataToServer "Update blending_header set status=1 where lotNo ='" & strLotNo & "'"
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   OpenDetail MyDDE.GetFieldByName("IDTrans")
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:
Select Case Index
       Case 0:
            RcLookup.DBOpen "select blending_header.noItem,blending_header.lotNo,blending_header.hasil_powder,inventory.itemName,inventory.UOM from blending_header inner join inventory on blending_header.noItem = inventory.noItem where blending_header.status=0", CNN, lckLockReadOnly
            
End Select
If RcLookup.Recordcount <> 0 Then
   Select Case Index
          Case 0:
            mCall.FromTagActive = "Produk Dikirim"
          Case 1:
            mCall.FromTagActive = "Gudang Tujuan"
   End Select
   Set mCall.FormData = RcLookup.DBRecordset
   mCall.LookUp Me

Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgCrtical
   If MyDDE.ChildRecordset.Recordcount <> 0 Then
      MyDDE.ChildRecordset.CancelBatch adAffectCurrent
      If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
   End If
End If
'
Exit Sub
Hell:
    Err.Clear
End Sub

Private Sub OpenDetail(ByVal ParameterString As String)
   If ParameterString = "" Then ParameterString = "xxxxxxxx"
   RcDetail.DBOpen "SELECT backflush_output.noItem,inventory.itemName,backflush_output.sl_no as ln_no,backflush_output.output_qty as kuantitas,inventory.UOM as satuan,backflush_output.Description as keterangan FROM backflush_output inner join inventory on backflush_output.noItem = inventory.noItem where IDTrans='" & ParameterString & "'", CNN, lckLockBatch
   Set MyDDE.ChildRecordset = RcDetail.DBRecordset  '.Clone(adLockBatchOptimistic)
   Set DGPurchase.DataSource = MyDDE.ChildRecordset
End Sub

Private Sub SimpanDetail()
   With MyDDE.ChildRecordset
        If .Recordcount <> 0 Then
              If SendDataToServer("delete from backflush_output where idTrans='" & MyDDE.GetFieldByName("IDTrans") & "'") = True Then
                  .MoveFirst
                      While Not .EOF
                         SendDataToServer "insert into backflush_output (idx,IDTrans,NoItem,sl_no,output_qty,Description,OrderID) values (newID(),'" & _
                                          MyDDE.GetFieldByName("IDTrans") & "','" & .Fields("NoItem") & "','" & .Fields("Ln_No") & "'," & _
                                          FQty(.Fields("kuantitas")) & ",'" & .Fields("keterangan") & "','-')"
                         UpdateBlending .Fields("Ln_No")
                         .MoveNext
                      Wend
                  .MoveLast
                  DGPurchase.Refresh
              End If
        End If
   End With
End Sub


Private Sub PrepareQuery()
   With MyDDE
       
       .PrepareAppend = "insert into backflush_header (IDTrans,DateTrans,[issued by],typetrans) values ('" & txtBox(0).Text & "','" & _
                         Format(DTPicker1.Value, "yyyy-MM-dd") & "','" & MainMenu.StatusBar1.Panels(1).Text & "','FG')"
                        
       
       .PrepareUpdate = "update backflush_header set DateTrans='" & Format(DTPicker1.Value, "yyyy-MM-dd") & "',[issued by]='" & _
                        MainMenu.StatusBar1.Panels(1).Text & "' where IDTrans='" & .GetFieldByName("IDTrans") & "'"
                        
       .PrepareDelete = "delete from backflush_header WHERE IDTrans = '" & .GetFieldByName("IDTrans") & "'"
   End With
End Sub






