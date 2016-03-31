VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmConfigRL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Konfigurasi RL Batch"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfigRL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   9870
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   5490
      Left            =   0
      ScaleHeight     =   5490
      ScaleWidth      =   9870
      TabIndex        =   1
      Top             =   0
      Width           =   9870
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   4320
         Picture         =   "frmConfigRL.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   510
         Width           =   330
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "itemName"
         Height          =   315
         Index           =   2
         Left            =   1245
         TabIndex        =   8
         Tag             =   "RL"
         Top             =   510
         Width           =   3075
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "qty_in"
         Height          =   315
         Index           =   1
         Left            =   6510
         Locked          =   -1  'True
         TabIndex        =   6
         Tag             =   "RL"
         Top             =   195
         Width           =   960
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "sl_no"
         Height          =   315
         Index           =   0
         Left            =   1245
         TabIndex        =   5
         Tag             =   "RL"
         Top             =   180
         Width           =   3405
      End
      Begin MSDataGridLib.DataGrid gridRL 
         Height          =   4425
         Left            =   60
         TabIndex        =   2
         Tag             =   "RL"
         Top             =   960
         Width           =   9750
         _ExtentX        =   17198
         _ExtentY        =   7805
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   2
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "transID"
            Caption         =   "No Penerimaan"
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
            DataField       =   "supplier"
            Caption         =   "Supplier"
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
            DataField       =   "qty"
            Caption         =   "Qty"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column02 
            EndProperty
         EndProperty
      End
      Begin VB.Line Line3 
         X1              =   6600
         X2              =   5460
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line2 
         X1              =   1395
         X2              =   180
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Line Line1 
         X1              =   1350
         X2              =   165
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         Height          =   285
         Index           =   2
         Left            =   7530
         TabIndex        =   7
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   285
         Index           =   1
         Left            =   5460
         TabIndex        =   4
         Top             =   225
         Width           =   1470
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "No RL Batch"
         Height          =   285
         Index           =   0
         Left            =   195
         TabIndex        =   3
         Top             =   225
         Width           =   1470
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rumput Laut"
         Height          =   285
         Index           =   3
         Left            =   195
         TabIndex        =   9
         Top             =   555
         Width           =   1470
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5505
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   1005
      BindFormTAG     =   "RL"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmConfigRL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RsDetail As New DBQuick
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RcPartner As New DBQuick

Private Sub cmdLink_Click(Index As Integer)
   OpenPartner 0
End Sub

Private Sub Form_Load()
   HiasFormManTell Picture1, Me
   With MyDDE
       .EditModeReplace = False
       Set .BindForm = frmConfigRL 'ME
       .BindFormTAG = "RL"
       Set .ActiveConnection = CNN
       '.PrepareQuery = "select [inventory tabel].sl_no,[inventory tabel].qty_in,[inventory tabel].noItem,[inventory tabel].lokasiGdg,inventory.itemName from [inventory tabel] inner join inventory on [inventory tabel].noItem = inventory.noItem where lockFIFO=0 and left([inventory tabel].noItem,5)='BB-BA'"
    
       .PrepareQuery = "SELECT dbo.[Inventory Tabel].sl_no, dbo.[Inventory Tabel].QTY_IN, dbo.[Inventory Tabel].NoItem, dbo.[Inventory Tabel].LokasiGdg, " & _
                               " dbo.Inventory.ItemName " & _
                               " FROM  dbo.[Inventory Tabel] INNER JOIN " & _
                               " dbo.Inventory ON dbo.[Inventory Tabel].NoItem = dbo.Inventory.NoItem INNER JOIN " & _
                               " dbo.rl_config ON dbo.[Inventory Tabel].sl_no = dbo.rl_config.no_rl " & _
                               " WHERE (dbo.[Inventory Tabel].LockFIFO = 0) AND (LEFT(dbo.[Inventory Tabel].NoItem, 5) = 'BB-BA')"
   
   End With
   Set mCall = New frmCaller
   gridRL.HeadLines = 2

End Sub

Private Sub loadDetail()
  RsDetail.DBOpen "select * from view_rl_config where no_rl='" & MyDDE.GetFieldByName("sl_no") & "'", CNN, lckLockBatch
  Set MyDDE.ChildRecordset = RsDetail.DBRecordset
  Set gridRL.DataSource = MyDDE.ChildRecordset
End Sub


Private Sub mCall_BeforeUnload()
On Error Resume Next
Dim jml As Double
   Select Case mCall.FromTagActive
      Case "Daftar Penerimaan Rumput Laut"
         jml = 0
         If MyDDE.ChildRecordset.Recordcount > 0 Then
            MyDDE.ChildRecordset.MoveFirst
            While Not MyDDE.ChildRecordset.EOF
               jml = jml + Val(MyDDE.ChildRecordset.Fields("qty"))
               MyDDE.ChildRecordset.MoveNext
            Wend
         End If
         txt(1).Text = jml
   End Select
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
   Select Case mCall.FromTagActive
      Case "Daftar Rumput Laut"
         MyDDE.GetFieldByName("noItem") = mCall.GetFieldByName("noItem")
         MyDDE.GetFieldByName("lokasiGdg") = mCall.GetFieldByName("warehouse")
         MyDDE.GetFieldByName("itemName") = mCall.GetFieldByName("ItemName")
         
      Case "Daftar Penerimaan Rumput Laut"
         MyDDE.ChildRecordset.Fields("supplier") = mCall.GetFieldByName("supplier")
         MyDDE.ChildRecordset.Fields("TRansID") = mCall.GetFieldByName("no penerimaan")
         MyDDE.ChildRecordset.Fields("qty") = mCall.GetFieldByName("qty")
         MyDDE.ChildRecordset.Fields("noIdx") = mCall.GetFieldByName("noIdx")
         
   End Select
'   gridRL.SetFocus
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
   cmdLink(0).Enabled = False
   Select Case AdReasonActiveDb
      Case tmbAddNew
         Dim IDGen As New IDGenerator
         cmdLink(0).Enabled = True
         MyDDE.GetFieldByName("sl_no") = IDGen.GetID("RLB")
         MyDDE.GetFieldByName("qty_in") = 0
         
      Case tmbEdit
         cmdLink(0).Enabled = True
      Case tmbDetail
         OpenPartner 1
      Case tmbSave
         SimpanDetail
      Case tmbDelete
         DeleteDetail
   End Select
End Sub

Private Sub SimpanDetail()
   SendDataToServer "delete from rl_config where no_rl='" & MyDDE.GetFieldByName("sl_no") & "'"
   With MyDDE.ChildRecordset
      .MoveFirst
      While Not .EOF
          SendDataToServer "insert into rl_config(no_rl,TransID,qty,noIdx) values ('" & txt(0) & "','" & .Fields("TransID") & "'," & .Fields("qty") & ",'" & .Fields("noIdx") & "')"
          SendDataToServer "update [Inventory tabel] set QTy_out = Qty_out + " & .Fields("Qty") & ", StockTmp = stockTmp - " & .Fields("Qty") & " where noIdx ='" & .Fields("noIdx") & "'"
         .MoveNext
      Wend
   End With
End Sub


Private Sub DeleteDetail()
On Error Resume Next
    'SendDataToServer "delete from rl_config where no_rl='" & MyDDE.GetFieldByName("sl_no") & "'"
   With MyDDE.ChildRecordset
      .MoveFirst
       While Not .EOF
          SendDataToServer "update [inventory tabel] set qty_in = qty_in + " & .Fields("Qty") & ", StockTmp = stockTmp + " & .Fields("Qty") & " where noIdx='" & .Fields("noIdx") & "'"
         .MoveNext
      Wend
   End With
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   If MyDDE.ActiveRecordset.Recordcount > 0 Then
      loadDetail
   End If
   'Set gridRL.DataSource = MyDDE.ChildRecordset
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbDelete:
               PrepareQuery
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               PrepareQuery
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbCancel:
            MyDDE.IsChildMemberReady = False
End Select
End Sub

Private Sub PrepareQuery()
   MyDDE.PrepareAppend = "INSERT INTO [Inventory Tabel]([NoItem] " & _
                                 ",[QTY_IN]" & _
                                 ",[PriceIn],[QTY_OUT],[PriceOut],[RefTrans]" & _
                                 ",[DateTrans]" & _
                                 ",[DateIssued]" & _
                                 ",[StockTmp],[LockFIFO],[TypeTrans],[HPP],[QTY ADJ],[QTY Real],[DateAdj]" & _
                                 ",[LokasiGdg]" & _
                                 ",[sl_no]) " & _
                          " Values ('" & MyDDE.GetFieldByName("noItem") & "'" & _
                                 "," & txt(1) & _
                                 ",0,0,0,''" & _
                                 ",'" & Format(Now, "yyyy-MM-dd") & "'" & _
                                 ",'" & Format(Now, "yyyy-MM-dd") & "'" & _
                                 ", " & txt(1) & _
                                 ",0,'AP',0,0,0" & _
                                 ",'" & Format(Now, "yyyy-MM-dd") & "'" & _
                                 ",'" & MyDDE.GetFieldByName("lokasiGdg") & "'" & _
                                 ",'" & txt(0) & "')"
   
   MyDDE.PrepareDelete = "delete from [inventory tabel] where sl_no='" & MyDDE.GetFieldByName("sl_no") & "'"
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:
Select Case Index
       Case 0: 'Order normal
            RcPartner.DBOpen "select noItem,itemName,warehouse from inventory where left(noItem,5)='BB-BA'", CNN, lckLockReadOnly
       Case 1: 'Detail Order SPP
            RcPartner.DBOpen "select * from view_lookup_rl_batch ", CNN, lckLockReadOnly
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 0:
            mCall.FromTagActive = "Daftar Rumput Laut"
            mCall.CaptionLink = "Daftar Rumput Laut"
          Case 1:
            mCall.FromTagActive = "Daftar Penerimaan Rumput Laut"
            mCall.CaptionLink = "Daftar Penerimaan Rumput Laut"
   End Select
   'Set mCall = New frmCaller
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgCrtical
End If
'
Exit Sub
Hell:
    Err.Clear
End Sub

