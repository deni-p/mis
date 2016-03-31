VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{03B0319E-A65B-4284-85CA-DAE87D7F78DA}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmInvoice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Pembelian"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmInvoice.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   10185
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5595
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   1005
      BindFormTAG     =   "inv"
      ActiveLanguage  =   1
   End
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
      Height          =   5610
      Left            =   0
      ScaleHeight     =   5610
      ScaleWidth      =   10185
      TabIndex        =   17
      Top             =   0
      Width           =   10185
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "TermPayment"
         Height          =   330
         Index           =   1
         Left            =   5910
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "PO"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "TOTAL_PPN"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   14
         Tag             =   "INV"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "invoice_no"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   1
         Tag             =   "INV"
         Top             =   120
         Width           =   1935
      End
      Begin VB.TextBox txt 
         DataField       =   "status"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   5915
         Locked          =   -1  'True
         TabIndex        =   13
         Tag             =   "INV"
         Top             =   1935
         Width           =   1935
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "partner_id"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   330
         Index           =   4
         Left            =   1680
         TabIndex        =   3
         Tag             =   "INV"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "distributor"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   330
         Index           =   5
         Left            =   2805
         TabIndex        =   4
         Tag             =   "INV"
         Top             =   480
         Width           =   4695
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "id_cur"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   330
         Index           =   6
         Left            =   1680
         TabIndex        =   6
         Tag             =   "INV"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "currName"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   330
         Index           =   7
         Left            =   2805
         TabIndex        =   7
         Tag             =   "INV"
         Top             =   840
         Width           =   4695
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "in_charge"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   330
         Index           =   8
         Left            =   1680
         TabIndex        =   9
         Tag             =   "INV"
         Top             =   1200
         Width           =   6165
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "no_invoice"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   330
         Index           =   9
         Left            =   1680
         TabIndex        =   10
         Tag             =   "INV"
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   7500
         MaskColor       =   &H000000C0&
         Picture         =   "FrmInvoice.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "SPPH"
         Top             =   488
         UseMaskColor    =   -1  'True
         Width           =   350
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   7500
         MaskColor       =   &H000000C0&
         Picture         =   "FrmInvoice.frx":6BDC
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "SPPH"
         Top             =   848
         UseMaskColor    =   -1  'True
         Width           =   350
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "jumlah_pembayaran"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   330
         Index           =   15
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   12
         Tag             =   "INV"
         Top             =   1920
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTP 
         DataField       =   "tanggal"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   0
         Left            =   5915
         TabIndex        =   2
         Tag             =   "INV"
         Top             =   105
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
         Format          =   58589187
         CurrentDate     =   39393
      End
      Begin MSComctlLib.ListView listInv 
         Height          =   2775
         Left            =   75
         TabIndex        =   16
         Tag             =   "INV"
         Top             =   2760
         Width           =   9945
         _ExtentX        =   17542
         _ExtentY        =   4895
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No PO"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Kode Brg"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nama Barang"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Satuan"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Qty"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Harga"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Diskon"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "PPN"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Total Harga"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTP 
         DataField       =   "tgl_invoice"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   1
         Left            =   5915
         TabIndex        =   11
         Tag             =   "INV"
         Top             =   1568
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
         Format          =   58589187
         CurrentDate     =   39393
      End
      Begin VB.Line Line1 
         Index           =   15
         X1              =   4200
         X2              =   6090
         Y1              =   2595
         Y2              =   2595
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Termin Pembayaran"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   5
         Left            =   4230
         TabIndex        =   28
         Top             =   2348
         Width           =   1425
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   120
         X2              =   2040
         Y1              =   2595
         Y2              =   2595
      End
      Begin VB.Label LBLIdent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total PPn"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   27
         Top             =   2355
         Width           =   675
      End
      Begin VB.Label LBLIdent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   4230
         TabIndex        =   26
         Top             =   1988
         Width           =   465
      End
      Begin VB.Label LBLIdent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kurs"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   64
         Left            =   150
         TabIndex        =   25
         Top             =   908
         Width           =   315
      End
      Begin VB.Label LBLIdent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "In Charge"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   62
         Left            =   150
         TabIndex        =   24
         Top             =   1268
         Width           =   720
      End
      Begin VB.Label LBLIdent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   61
         Left            =   150
         TabIndex        =   23
         Top             =   188
         Width           =   465
      End
      Begin VB.Label LBLIdent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   60
         Left            =   4230
         TabIndex        =   22
         Top             =   165
         Width           =   570
      End
      Begin VB.Label LBLIdent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   21
         Top             =   555
         Width           =   570
      End
      Begin VB.Label LBLIdent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Tagihan"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   20
         Top             =   1988
         Width           =   975
      End
      Begin VB.Label LBLIdent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Invoice"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   19
         Top             =   1628
         Width           =   825
      End
      Begin VB.Label LBLIdent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Invoice"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   8
         Left            =   4230
         TabIndex        =   18
         Top             =   1628
         Width           =   1140
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   120
         X2              =   2040
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   120
         X2              =   2040
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   120
         X2              =   2040
         Y1              =   1155
         Y2              =   1155
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   120
         X2              =   2040
         Y1              =   1515
         Y2              =   1515
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   120
         X2              =   2040
         Y1              =   1875
         Y2              =   1875
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   120
         X2              =   2040
         Y1              =   2235
         Y2              =   2235
      End
      Begin VB.Line Line1 
         Index           =   9
         X1              =   4200
         X2              =   7080
         Y1              =   405
         Y2              =   405
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   4200
         X2              =   6840
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Line Line1 
         Index           =   11
         X1              =   4200
         X2              =   6120
         Y1              =   1868
         Y2              =   1868
      End
   End
End
Attribute VB_Name = "FrmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Dim rsPop As New DBQuick
Dim IDGen As New IDGenerator
Dim isEdit As Boolean
Dim isAppend As Boolean
Dim MyJournal As New clsJournal
Dim Cancelled As Boolean
Dim LastID As String

Private Sub cmdLink_Click(Index As Integer)
   OpenPartner Index
End Sub

Private Sub DDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo xErr
   cmdLink(0).Enabled = False
   cmdLink(1).Enabled = False
   Select Case AdReasonActiveDb
      Case tmbAddNew:
         cmdLink(0).Enabled = True
         cmdLink(1).Enabled = True
         DDE.GetFieldByName("invoice_no") = IDGen.GetID("IS")
         DTP(0).Value = Now
         DTP(1).Value = Now
         DDE.GetFieldByName("tanggal") = Now
         DDE.GetFieldByName("tgl_invoice") = Now
         DDE.GetFieldByName("total_tagihan") = 0
         DDE.GetFieldByName("in_charge") = MainMenu.StatusBar1.Panels(1).Text
         DDE.GetFieldByName("status") = "APPROVED"
         
      Case tmbSave:
         If DDE.IsChildMemberReady Then
            If isAppend Then
               SendDataToServer " INSERT INTO TransData (TransID,PartnerID ,Currid, DateTrans,  TypeTrans) " & _
                                 " VALUES ('" & DDE.GetFieldByName("invoice_no") & _
                                         "','" & DDE.GetFieldByName("partner_id") & _
                                         "','" & DDE.GetFieldByName("id_cur") & _
                                         "','" & Format(DDE.GetFieldByName("tanggal"), "yyyy-MM-dd") & _
                                         "','SI')"
            ElseIf isEdit Then
               SendDataToServer " update transData set partnerID='" & DDE.GetFieldByName("invoice_no") & _
                                                        "',currid='" & DDE.GetFieldByName("partner_id") & _
                                                        "',dateTrans='" & DDE.GetFieldByName("id_cur") & _
                                                        "' where TransID ='" & DDE.GetFieldByName("invoice_no") & "'"
            End If
            
             
            SaveDetail
            isEdit = False
            isAppend = False
         End If
         
      Case tmbDelete:
         DelDetail
         
   End Select
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub SaveDetail()
Dim X, Y As Integer
Dim strSQL As String
Dim JmlTotal As Double
Dim TmpPo As String
On Error GoTo xErr
   SendDataToServer "delete from invoicing_line where invoice_no ='" & DDE.GetFieldByName("invoice_no") & "'"
   SendDataToServer "delete from [detail transData] where transID='" & DDE.GetFieldByName("invoice_no") & "'"
   Y = 1
   JmlTotal = 0
   TmpPo = listInv.ListItems(1).Text
   For X = 1 To listInv.ListItems.Count
      If listInv.ListItems(X).Checked Then
         strSQL = "insert into invoicing_line (invoice_no,PurchaseID,total_tagihan,NoItem ,qty,harga,total,total_ppn,discount,total_sisa,total_pph,no_urut) values ('" & _
                    DDE.GetFieldByName("invoice_no") & _
                "' ,'" & listInv.ListItems(X).Text & _
                "', " & FQty(listInv.ListItems(X).SubItems(8)) & _
                " ,'" & listInv.ListItems(X).SubItems(1) & _
                "', " & FQty(listInv.ListItems(X).SubItems(4)) & _
                " , " & FQty(listInv.ListItems(X).SubItems(5)) & _
                " , " & FQty(listInv.ListItems(X).SubItems(8)) & _
                " , " & FQty(listInv.ListItems(X).SubItems(6)) & _
                " , " & FQty(listInv.ListItems(X).SubItems(7)) & ",0,0," & X & ")"
         
         SendDataToServer strSQL
         
         SendDataToServer "insert into [detail transData] (transID,noItem,DateTrans,Qty_In,Qty_receive,price) values ('" & _
                DDE.GetFieldByName("invoice_no") & _
                "','" & listInv.ListItems(X).SubItems(1) & _
                "','" & Format(DDE.GetFieldByName("tanggal"), "yyyy-MM-dd") & _
                "', " & FQty(listInv.ListItems(X).SubItems(4)) & _
                " , " & FQty(listInv.ListItems(X).SubItems(4)) & _
                " , " & FQty(listInv.ListItems(X).SubItems(5)) & ")"
         
         
         strSQL = "update [detail PO] set statusTrans=6 where noItem='" & listInv.ListItems(X).SubItems(1) & "' and PurchaseID='" & listInv.ListItems(X).Text & "'"
         SendDataToServer strSQL
         
         UpdateReceiveNote listInv.ListItems(X).Text, listInv.ListItems(X).SubItems(1)
         
         JmlTotal = JmlTotal + FQty(listInv.ListItems(X).SubItems(8))
         If (TmpPo <> listInv.ListItems(X).Text) Or (X = listInv.ListItems.Count) Then
            'kirim data ke voucher
            SendVoucher TmpPo, txt(4), "-", DTP(0).Value, JmlTotal, 0, txt(9), "AP"
            TmpPo = listInv.ListItems(X).Text
            JmlTotal = 0
         End If
         Y = Y + 1
      End If
   Next
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear
End Sub

Private Sub UpdateReceiveNote(pNoPO As String, pNoItem As String)
   Dim rsDN As New DBQuick
   Dim aDn As String
   rsDN.DBOpen "select DNID from [detail transdata] where transID='" & pNoPO & "' and noItem='" & pNoItem & "'", CNN
   If rsDN.DBRecordset.Recordcount > 0 Then
      aDn = rsDN.DBRecordset.Fields(0)
      SendDataToServer "update [detail transdata] set status=1,statusInvoice=1"
   End If
   rsDN.CloseDB
End Sub

Private Sub DelDetail()
On Error GoTo xErr
   SendDataToServer "delete from invoicing_line where invoice_no ='" & DDE.GetFieldByName("No_invoice") & "'"
   SendDataToServer "delete from [detail transData] where transID ='" & DDE.GetFieldByName("invoice_no") & "'"
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear
End Sub

Private Sub DDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   showDetail
End Sub

Private Sub DDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim IsSelected As Boolean
Dim X As Integer
   Cancelled = False
   IsSelected = False
   Select Case AdReasonActiveDb
      Case tmbAddNew:
         isAppend = True
         LastID = txt(1).Text
      Case tmbEdit:
         isEdit = True
         LastID = txt(1).Text
      Case tmbSave
         If listInv.ListItems.Count > 0 Then
            For X = 1 To listInv.ListItems.Count
               If listInv.ListItems(X).Checked Then IsSelected = True
            Next
            If IsSelected Then
               prepareSQL
               DDE.IsChildMemberReady = True
            Else
               MessageBox "Belum ada item yang dipilih", "Peringatan", msgOkOnly, msgCrtical
               DDE.IsChildMemberReady = False
            End If
         Else
            MessageBox " Item Data tidak tersedia ", "Peringatan", msgOkOnly, msgCrtical
            DDE.IsChildMemberReady = False
         End If
      Case tmbCancel:
         
         Cancelled = True
         
      Case tmbDelete
         prepareSQL
   End Select
End Sub


Private Sub Form_Load()
   'HiasForm Picture1, Me
   HiasFormManTell Picture2, Me
   Set mCall = New frmCaller
   Set DDE.BindForm = Me
   Set DDE.ActiveConnection = CNN
   DDE.PrepareQuery = "select * from invoice_header"
   DDE.SetPermissions = aksess.MayDo("Invoice Pembelian")
End Sub

Private Sub showDetail()
   Dim RsDetail As New DBQuick
   Dim vLst As ListItem
   Dim X As Integer
   Dim AllTotal As Double
   
   If IsEmpty(DDE.GetFieldByName("invoice_no")) Then
      RsDetail.DBOpen "Select * from invoice_detail where invoice_no=''", CNN
   Else
      If Not Cancelled Then
         txt(1).Text = DDE.GetFieldByName("invoice_no")
         RsDetail.DBOpen "Select * from invoice_detail where invoice_no='" & IIf(IsNull(DDE.GetFieldByName("invoice_no")), 0, DDE.GetFieldByName("invoice_no")) & "'", CNN
      Else
         RsDetail.DBOpen "Select * from invoice_detail where invoice_no='" & LastID & "'", CNN
      End If
   End If
   
   listInv.ListItems.Clear
   AllTotal = 0
   If RsDetail.Recordcount > 0 Then
      RsDetail.DBRecordset.MoveFirst
      For X = 1 To RsDetail.DBRecordset.Recordcount
         Set vLst = listInv.ListItems.Add(X, "A" & X, RsDetail.DBRecordset.Fields("PurchaseID"))
         vLst.Checked = True
         listInv.ListItems(X).ListSubItems.Add 1, "KD", IIf(IsNull(RsDetail.DBRecordset.Fields("NoItem")), "", RsDetail.DBRecordset.Fields("NoItem"))
         listInv.ListItems(X).ListSubItems.Add 2, "BRG", IIf(IsNull(RsDetail.DBRecordset.Fields("NamaBarang")), "", RsDetail.DBRecordset.Fields("NamaBarang"))
         listInv.ListItems(X).ListSubItems.Add 3, "UOM", IIf(IsNull(RsDetail.DBRecordset.Fields("Satuan")), "", RsDetail.DBRecordset.Fields("Satuan"))
         listInv.ListItems(X).ListSubItems.Add 4, "Qty", Format(RsDetail.DBRecordset.Fields("Qty"), QtyForm)
         listInv.ListItems(X).ListSubItems.Add 5, "HRG", Format(CDbl(RsDetail.DBRecordset.Fields("Harga")), QtyForm)
         listInv.ListItems(X).ListSubItems.Add 6, "DIS", Format(CDbl(RsDetail.DBRecordset.Fields("Discount")), QtyForm)
         listInv.ListItems(X).ListSubItems.Add 7, "PPN", Format(CDbl(RsDetail.DBRecordset.Fields("total_ppn")), QtyForm)
         listInv.ListItems(X).ListSubItems.Add 8, "TOT", Format(CDbl(RsDetail.DBRecordset.Fields("Total")), QtyForm)
         AllTotal = AllTotal + CDbl(RsDetail.DBRecordset.Fields("Total"))
         RsDetail.DBRecordset.MoveNext
      Next
      RsDetail.DBRecordset.MoveFirst
   End If
'   If Not Cancelled Then txt(15).Text = AllTotal
End Sub

Private Sub OpenPartner(aID As Integer)

Select Case aID
   Case 0
      rsPop.DBOpen "select * from supplier_invoice_pembelian", CNN
   Case 1
      rsPop.DBOpen "select * from [Currency Setup]", CNN
End Select
If rsPop.Recordcount <> 0 Then
    Select Case aID
       Case 0
          mCall.CaptionLink = "Supplier"
          mCall.FromTagActive = "Supplier"
          
       Case 1
          mCall.FromTagActive = "Currency"
          mCall.CaptionLink = "Currency"
    End Select
    Set mCall.FormData = rsPop.DBRecordset
    mCall.LookUp Me
Else
   MessageBox "Data belum ada atau data masih kosong..", "Invoice Pembelian", msgOkOnly, msgExclamation
End If

End Sub

Private Sub listInv_ItemCheck(ByVal Item As MSComctlLib.ListItem)
   Dim X As Integer
   Dim totInv As Double
   Dim totPPn As Double
'   txt(15) = "0"
'   txt(0) = "0"
   totInv = 0
   For X = 1 To listInv.ListItems.Count
      If listInv.ListItems(X).Checked Then
         totInv = totInv + listInv.ListItems(X).SubItems(8)
         totPPn = totPPn + listInv.ListItems(X).SubItems(7)
      End If
   Next
   txt(15) = totInv
   txt(0) = totPPn
End Sub

Private Sub mCall_BeforeUnload()
   Select Case mCall.FromTagActive
      Case "Currency":
         LoadItemBarang DDE.GetFieldByName("partner_id"), DDE.GetFieldByName("ID_cur")
   End Select

End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
   Select Case UCase(TagForm)
      Case "SUPPLIER":
         DDE.GetFieldByName("partner_id") = mCall.GetFieldByName("Kode")
         DDE.GetFieldByName("Distributor") = mCall.GetFieldByName("Supplier")
         txtBox(1).Text = mCall.GetFieldByName("termin")
      Case "CURRENCY":
         DDE.GetFieldByName("ID_cur") = mCall.GetFieldByName("CurrID")
         DDE.GetFieldByName("CurrName") = mCall.GetFieldByName("Currency Name")
         LoadItemBarang DDE.GetFieldByName("partner_id"), DDE.GetFieldByName("ID_cur")

   End Select
End Sub

Private Sub LoadItemBarang(aSupplier As String, aCurrency As String)
   Dim rsItemBrg As New DBQuick
   Dim X As Integer
   Dim vLst As ListItem
   Dim lDisc As Double
   Dim lPPN As Double
   
   If aSupplier = "" Then aSupplier = "xxxxxxx"
   rsItemBrg.DBOpen "select * from itemInvoice where partnerID='" & aSupplier & "' and CurID ='" & aCurrency & "' order by purchaseID", CNN
   listInv.ListItems.Clear
   If rsItemBrg.Recordcount > 0 Then
      rsItemBrg.DBRecordset.MoveFirst
      For X = 1 To rsItemBrg.DBRecordset.Recordcount
         Set vLst = listInv.ListItems.Add(X, "A" & X, rsItemBrg.DBRecordset.Fields("PurchaseID"))
         vLst.Checked = False
         listInv.ListItems(X).ListSubItems.Add 1, "KD", IIf(IsNull(rsItemBrg.DBRecordset.Fields("NoItem")), "", rsItemBrg.DBRecordset.Fields("NoItem"))
         listInv.ListItems(X).ListSubItems.Add 2, "BRG", IIf(IsNull(rsItemBrg.DBRecordset.Fields("ItemName")), "", rsItemBrg.DBRecordset.Fields("ItemName"))
         listInv.ListItems(X).ListSubItems.Add 3, "UOM", IIf(IsNull(rsItemBrg.DBRecordset.Fields("uom")), "", rsItemBrg.DBRecordset.Fields("uom"))
         listInv.ListItems(X).ListSubItems.Add 4, "Qty", Format(rsItemBrg.DBRecordset.Fields("QtyPO"), QtyForm)
         listInv.ListItems(X).ListSubItems.Add 5, "HRG", Format(rsItemBrg.DBRecordset.Fields("POPrice"), QtyForm)
         lDisc = CDbl(rsItemBrg.DBRecordset.Fields("Discount")) * CDbl(rsItemBrg.DBRecordset.Fields("QtyPO")) * CDbl(rsItemBrg.DBRecordset.Fields("POPrice")) / 100
         listInv.ListItems(X).ListSubItems.Add 6, "DIS", Format(lDisc, QtyForm)
         lPPN = CDbl(rsItemBrg.DBRecordset.Fields("VAT")) * ((CDbl(rsItemBrg.DBRecordset.Fields("QtyPO")) * CDbl(rsItemBrg.DBRecordset.Fields("POPrice"))) - lDisc) / 100
         listInv.ListItems(X).ListSubItems.Add 7, "PPN", Format(lPPN, QtyForm)
         listInv.ListItems(X).ListSubItems.Add 8, "TOT", Format((CDbl(rsItemBrg.DBRecordset.Fields("QtyPO")) * CDbl(rsItemBrg.DBRecordset.Fields("POPrice"))) - lDisc + lPPN, QtyForm)
         rsItemBrg.DBRecordset.MoveNext
      Next
      rsItemBrg.DBRecordset.MoveFirst
   End If
   
   Set rsItemBrg = Nothing
End Sub

Private Sub prepareSQL()
   Dim strSQL As String

On Error GoTo xErr
   
   strSQL = "insert into invoicing_header (invoice_no,tanggal,ID_Cur,partner_id,tgl_invoice,no_invoice,total_tagihan,in_charge,status,jumlah_pembayaran,total_ppn,total_pph) values ('" & _
                                               DDE.GetFieldByName("invoice_no") & _
                                        "','" & Format(DDE.GetFieldByName("tanggal"), "yyyy-MM-dd") & _
                                        "','" & DDE.GetFieldByName("ID_CUR") & _
                                        "','" & DDE.GetFieldByName("partner_id") & _
                                        "','" & Format(DDE.GetFieldByName("tgl_invoice"), "yyyy-MM-dd") & _
                                        "','" & DDE.GetFieldByName("No_invoice") & _
                                        "', " & FQty(txt(15).Text) & _
                                        " ,'" & DDE.GetFieldByName("in_charge") & _
                                        "','" & DDE.GetFieldByName("status") & _
                                        "', " & FQty(txt(15).Text) & _
                                        " , " & FQty(txt(0).Text) & ",0)"
   Debug.Print strSQL
   DDE.PrepareAppend = strSQL
   DDE.PrepareUpdate = "update invoicing_header set tanggal ='" & Format(DDE.GetFieldByName("tanggal"), "yyyy-MM-dd") & _
                                        "',id_cur='" & DDE.GetFieldByName("ID_CUR") & _
                                        "',partner_id ='" & DDE.GetFieldByName("partner_id") & _
                                        "',tgl_invoice='" & Format(DDE.GetFieldByName("tgl_invoice"), "yyyy-MM-dd") & _
                                        "',no_invoice='" & DDE.GetFieldByName("No_invoice") & _
                                        "',total_tagihan=" & FQty(DDE.GetFieldByName("total_tagihan")) & _
                                        "',in_charge='" & DDE.GetFieldByName("in_charge") & _
                                        "',status='" & DDE.GetFieldByName("status") & "'" & _
                                        " ,jumlah_pembayaran=" & FQty(txt(15)) & ",total_ppn=" & _
                                        FQty(txt(0)) & " where invoice_no = '" & DDE.GetFieldByName("invoice_no") & "'"
                                        
   DDE.PrepareDelete = "delete from invoicing_header where invoice_no ='" & DDE.GetFieldByName("invoice_no") & "'"

Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

