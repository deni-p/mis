VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmMatRequest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Permintaan Barang"
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   1710
   ClientWidth     =   10590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMatRequest.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   Tag             =   "Material Requisition"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4590
      Left            =   0
      ScaleHeight     =   4590
      ScaleWidth      =   10590
      TabIndex        =   9
      Top             =   0
      Width           =   10590
      Begin VB.TextBox txtApproved 
         Appearance      =   0  'Flat
         DataSource      =   "MyDDE"
         Height          =   330
         Left            =   5730
         Locked          =   -1  'True
         MaxLength       =   249
         TabIndex        =   7
         Top             =   4080
         Width           =   2505
      End
      Begin VB.TextBox txtIssued 
         Appearance      =   0  'Flat
         DataField       =   "Issued By"
         DataSource      =   "MyDDE"
         Height          =   330
         Left            =   1470
         Locked          =   -1  'True
         MaxLength       =   249
         TabIndex        =   6
         Tag             =   "RN"
         Top             =   4065
         Width           =   2235
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   4740
         Picture         =   "FrmMatRequest.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   458
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         DataField       =   "Note"
         Height          =   330
         Left            =   1470
         MaxLength       =   249
         TabIndex        =   5
         Tag             =   "RN"
         Top             =   3720
         Width           =   6765
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "DateTrans"
         Height          =   330
         Left            =   1470
         TabIndex        =   4
         Tag             =   "RN"
         Top             =   810
         Width           =   3630
         _ExtentX        =   6403
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
         Format          =   71630851
         CurrentDate     =   38538
      End
      Begin MSDataGridLib.DataGrid DGPurchase 
         Bindings        =   "FrmMatRequest.frx":6BDC
         Height          =   2265
         Left            =   75
         TabIndex        =   8
         Tag             =   "RN"
         Top             =   1275
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   3995
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "noItem"
            Caption         =   "Item ID"
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
            DataField       =   "InternalName"
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
            DataField       =   "Description"
            Caption         =   "Tgl Kebutuhan"
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
            DataField       =   "penggunaan"
            Caption         =   "Penggunaan"
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
            DataField       =   "Qty Required"
            Caption         =   "Qty"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00;(#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "UOM"
            Caption         =   "Satuan"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "#,##0;(#,##0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "batch_lot"
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
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column06 
               Button          =   -1  'True
               Locked          =   -1  'True
            EndProperty
         EndProperty
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   105
         X2              =   1470
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Label lblSupplier 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "dept"
         DataSource      =   "MyDDE"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   1470
         TabIndex        =   2
         Top             =   450
         Width           =   3615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dept"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   0
         Left            =   105
         TabIndex        =   15
         Top             =   495
         Width           =   405
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   105
         X2              =   1470
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         DataField       =   "Tanggal"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   2
         Left            =   105
         TabIndex        =   14
         Top             =   870
         Width           =   645
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   105
         X2              =   1470
         Y1              =   4035
         Y2              =   4035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note"
         DataField       =   "Tanggal"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   3
         Left            =   105
         TabIndex        =   13
         Top             =   3765
         Width           =   405
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   105
         X2              =   1470
         Y1              =   4380
         Y2              =   4380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issued By"
         DataField       =   "Tanggal"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   5
         Left            =   105
         TabIndex        =   12
         Top             =   4110
         Width           =   780
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   4365
         X2              =   5730
         Y1              =   4395
         Y2              =   4395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approval By"
         DataField       =   "Tanggal"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   6
         Left            =   4365
         TabIndex        =   11
         Top             =   4110
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Trans"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   7
         Left            =   105
         TabIndex        =   10
         Top             =   165
         Width           =   735
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   105
         X2              =   1470
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "IDTrans"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1470
         TabIndex        =   1
         Tag             =   "RN"
         Top             =   105
         Width           =   3630
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   4590
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmMatRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As New DBQuick
Private RcPartner As New DBQuick
Private RcIssued As New DBQuick
Private RcApprov As New DBQuick
Private MyData As New clsTransaksi
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private mVarStage As String
Private lMode As String

Public Property Let SetTagForm(ByVal vNewValue As Variant)
       Me.Tag = ""
       Me.Tag = vNewValue
      HiasFormManTell Picture2, Me
End Property

Private Sub cmdLink_Click(Index As Integer)
   OpenPartner Index
End Sub



Private Sub Command1_Click()
   DGPurchase.Enabled = True
  If DGPurchase.Enabled Then MsgBox "A"
End Sub




Private Sub DGPurchase_ButtonClick(ByVal ColIndex As Integer)
   OpenPartner 5
End Sub

Private Sub Form_Load()
HiasFormManTell Picture2, Me
'OpenKaryawan
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmMatRequest
    .BindFormTAG = "RN"
    .SetPermissions = UserDeleteDenied
    Set .ActiveConnection = CNN
    If lMode = "production" Then
      If Me.Tag = "MATERIAL ISSUED" Then
         .PrepareQuery = "SELECT [backflush_Header].OrderID AS [Order ID], [Manufacture Order].OrderName AS [Nama Order], [Manufacture Order].Type AS [Job Type], [Manufacture Order].Status AS [Current Status], [backflush_Header].DateTrans, [backflush_Header].Note, [backflush_Header].[Issued BY], [backflush_Header].[Received By], [backflush_Header].IDTrans,[backflush_Header].Approved_by FROM [Manufacture Order] INNER JOIN [backflush_Header] ON [Manufacture Order].OrderID = [backflush_Header].OrderID WHERE ([backflush_Header].Status = 0) and ([issued By]='" & MainMenu.StatusBar1.Panels(1).Text & "') ORDER BY [backflush_Header].IDTrans, [backflush_Header].OrderID"
      Else
         .PrepareQuery = "SELECT [backflush_Header].OrderID AS [Order ID], [Manufacture Order].OrderName AS [Nama Order], [Manufacture Order].Type AS [Job Type], [Manufacture Order].Status AS [Current Status], [backflush_Header].DateTrans, [backflush_Header].Note, [backflush_Header].[Issued BY], [backflush_Header].[Received By], [backflush_Header].IDTrans,[backflush_Header].Approved_by FROM [Manufacture Order] INNER JOIN [backflush_Header] ON [Manufacture Order].OrderID = [backflush_Header].OrderID WHERE ([backflush_Header].Status = 1)  and ([issued By]='" & MainMenu.StatusBar1.Panels(1).Text & "') ORDER BY [backflush_Header].IDTrans, [backflush_Header].OrderID"
      End If
    Else
       .PrepareQuery = "select backflush_line.IDTrans,DateTrans,Note,[Issued By],[Received By],[backflush_Header].status,dept,approved_by from [backflush_Header] inner join backflush_line on backflush_header.IDTrans = backflush_line.IDtrans  group by backflush_line.IDTrans,DateTrans,Note,[Issued By],[Received By],[backflush_Header].status,dept,approved_by,[backflush_Header].typetrans having ([issued By]='" & MainMenu.StatusBar1.Panels(1).Text & "') and TypeTrans <> 'MI' and (sum(backflush_line.[Qty required]) > 0)"
    End If
   
    If MainMenu.Toolbar1.Buttons(5).Visible = True Then 'Menu Penjualan
        .SetPermissions = aksess.MayDo("Permintaan Barang")
    ElseIf MainMenu.Toolbar1.Buttons(7).Visible = True Then 'Menu Logistik
        .SetPermissions = aksess.MayDo("Permintaan Barang Logistik")
    ElseIf MainMenu.Toolbar1.Buttons(9).Visible = True Then  'Menu Gudang
        .SetPermissions = aksess.MayDo("Permintaan Barang Gudang")
    ElseIf MainMenu.Toolbar1.Buttons(13).Visible = True Then 'Menu ACCounting
        .SetPermissions = aksess.MayDo("Permintaan Barang Acc")
    End If
End With
Set mCall = New frmCaller
lblSupplier(0).ForeColor = vbWindowText
If MyDDE.ActiveRecordset.Recordcount > 0 Then MyDDE.ActiveRecordset.MoveFirst
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
rs.CloseDB
RcPartner.CloseDB
RcIssued.CloseDB
RcApprov.CloseDB
Set rs = Nothing
Set RcPartner = Nothing
Set RcIssued = Nothing
Set RcApprov = Nothing
Set MyData = Nothing
Set mCall = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmMatRequest = Nothing
End Sub

Private Sub mCall_BeforeUnload()
'Select Case mCall.FromTagActive
'       Case UCase(mCall.FromTagActive) = "STAGE":
'            If FindOwnRecordset(MyDDE.ChildRecordset, "[Item ID] = '" & MyDDE.ChildRecordset.Fields("Item ID") & "'") = True Then
'               MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Item ID") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
'               MyDDE.ChildRecordset.CancelBatch adAffectCurrent
'               If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
'            Else
'               'messagebox MyDDE.ChildRecordset.Fields("Item ID")
'               If IsNull(MyDDE.ChildRecordset.Fields("Item ID")) = True Or MyDDE.ChildRecordset.Fields("Item ID") = "" Then
'                  MyDDE.ChildRecordset.CancelBatch adAffectCurrent
'                  If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
'               End If
'            End If
'            If DGPurchase.Enabled = True Then
'               DGPurchase.AllowUpdate = True
'               DGPurchase.Col = 4
'               DGPurchase.SetFocus
'            End If
'End Select
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Dim Mrc As New DBQuick
Select Case TagForm
       Case "PRODUCTION ORDER":
            With MyDDE
                 .GetFieldByName(0) = mCall.GetFieldByName(0)
                 .GetFieldByName(1) = mCall.GetFieldByName(1)
                 .GetFieldByName(2) = mCall.GetFieldByName(2)
                 .GetFieldByName(3) = mCall.GetFieldByName(3)
            End With
            
       Case "Stage":
            With MyDDE.ChildRecordset
               .Fields("noItem") = mCall.GetFieldByName("kode")
               .Fields("InternalName") = mCall.GetFieldByName("Nama Barang")
               .Fields("UOM") = mCall.GetFieldByName("Satuan")
               .Fields("status") = 0
               .Fields("IDTrans") = Label2.Caption
            End With
            
       Case "Departemen":
            MyDDE.GetFieldByName("dept") = mCall.GetFieldByName("kode")
            lblSupplier(0).Caption = mCall.GetFieldByName("kode")
            
       Case "Manufacturing Order":
            MyDDE.GetFieldByName("Order ID") = mCall.GetFieldByName("Order ID")
            MyDDE.GetFieldByName("Job Type") = mCall.GetFieldByName("Job Type")
            MyDDE.GetFieldByName("Current Status") = mCall.GetFieldByName("Current Status")
      
      Case "Detail Komponen":
            Dim Rc As New DBQuick
            Rc.DBOpen "SELECT     [Ord Comp Detail].NoItem AS [Item ID], [Ord Comp Detail].[DESC] AS Description, [Ord Comp Detail].UOM, Inventory.WareHouse AS WareHouse, 0 AS Cost, [Ord Comp Detail].[Quote Qty],[Ord Comp Detail].[StageID] FROM [Ord Comp Detail] INNER JOIN Inventory ON [Ord Comp Detail].NoItem = Inventory.NoItem WHERE  ([Ord Comp Detail].OrderID = N'" & MyDDE.GetFieldByName(0) & "') order by  [Ord Comp Detail].NoItem", CNN
            With Rc.DBRecordset
                 If .Recordcount <> 0 Then
                    With MyDDE.ChildRecordset
                         If .Recordcount <> 0 Then
                          .MoveFirst
                          Do
                            If .EOF Then Exit Do
                            .Delete adAffectCurrent
                            .MoveNext
                          Loop
                         End If
                    End With
                    Do
                      If .EOF Then Exit Do
                        MyDDE.ChildRecordset.AddNew
                       With mCall
                        MyDDE.ChildRecordset.Fields(0) = .GetFieldByName(0)
                        MyDDE.ChildRecordset.Fields(1) = .GetFieldByName(1)
                        MyDDE.ChildRecordset.Fields(2) = .GetFieldByName(2)
                        MyDDE.ChildRecordset.Fields(3) = .GetFieldByName(3)
                        MyDDE.ChildRecordset.Fields("Qty WareHouse") = 0
                        MyDDE.ChildRecordset.Fields("Qty Received") = .GetFieldByName(5)
                        MyDDE.ChildRecordset.Fields("StageID") = .GetFieldByName("StageID")
                        MyDDE.ChildRecordset.Fields(7) = .GetFieldByName(7)
                        MyDDE.ChildRecordset.Fields("Cost") = .GetFieldByName("Cost")
                        'messagebox .GetFieldByName("Quote qty")
                        End With
                       .MoveNext
                    Loop
                    MyDDE.ChildRecordset.MoveFirst
                 End If
            End With
         
      Case "Batch No / Lot No"
            MyDDE.ChildRecordset.Fields("batch_lot") = mCall.GetFieldByName(0)
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim IDGen As New IDGenerator
cmdLink(2).Enabled = False
DGPurchase.AllowUpdate = False
Select Case AdReasonActiveDb
       Case tmbEdit:
         DGPurchase.AllowUpdate = True
         DGPurchase.Enabled = True
         cmdLink(2).Enabled = True
         'cmdLink(3).Enabled = True

         
       Case tmbAddNew:
            cmdLink(2).Enabled = True
            IDGen.ExtParameter = CurrentDept
            Label2 = IDGen.GetID("FR")
            MyDDE.GetFieldByName("Note") = "-"
        '   If CmdLink(2).Enabled = True Then CmdLink(2).SetFocus
            DTPicker1.Value = Now
            MyDDE.GetFieldByName("dateTrans") = Now
            MyDDE.GetFieldByName("dept") = IDGen.GetDept
            txtIssued.Text = MainMenu.StatusBar1.Panels(1).Text
            lblSupplier(0).Caption = NamaDept
            DGPurchase.AllowUpdate = True
       Case tmbDetail:
            If MyDDE.IsChildMemberReady = True Then
               
               If lMode = "production" Then
                  OpenPartner 4
               Else
                  OpenPartner 1
               End If
               MyDDE.ChildRecordset.Fields("description") = "-"
               MyDDE.ChildRecordset.Fields("penggunaan") = "-"
               MyDDE.ChildRecordset.Fields("Qty Required") = 1
               DGPurchase.AllowUpdate = True
               DGPurchase.Enabled = True
            End If
       Case tmbDelete:
            If MyDDE.IsChildMemberReady = True Then
               SendDataToServer "delete from backflush_line where IDTrans='" & Label2.Caption & "'"
            End If
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               With MyDDE.ChildRecordset
               If .Recordcount <> 0 Then
                   If SendDataToServer("DELETE FROM  [backflush_line] WHERE     (IDTrans = N'" & Label2 & "')") = True Then
                        .MoveFirst
                        Do
                          If .EOF Then Exit Do
                             SendDataToServer (" INSERT INTO  [backflush_line]" & _
                                               " (Idtrans, NoItem, Description, UOM, penggunaan, [Qty Required],status,batch_lot)" & _
                                               " VALUES  (N'" & Label2 & "',N'" & .Fields("noItem") & "', N'" & .Fields("Description") & "', N'" & .Fields("UOM") & "', N'" & .Fields("penggunaan") & "'," & FQty(.Fields("Qty Required")) & ",0,'" & .Fields("batch_lot") & "')")
                          .MoveNext
                        Loop
                        .MoveFirst
                    End If
               End If
               End With
               'cmdLink(0).Enabled = False
               DGPurchase.AllowUpdate = False
            End If
       Case tmbPrint:
            Dim aReport As New utility
            aReport.CallReportView "Select * from QuerySPB where idtrans='" & Label2.Caption & "'", "SuratPermintaanBarang.rpt", ReportPath, "Surat Permintaan Barang"
            Set aReport = Nothing
            
End Select
'cmdLink(0).Enabled = Text1.Enabled
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
Dim mVarStatus As Byte
With MyDDE
    If UCase(Me.Tag) = "MATERIAL ISSUED" Then
       mVarStatus = 0
    Else
       mVarStatus = 1
    End If
    If lMode = "production" Then
       .PrepareAppend = " INSERT INTO [backflush_Header]" & _
                        " (IDTrans, DateTrans, Note, [Issued BY], Status,dept,OrderID,typetrans)" & _
                        " VALUES  (N'" & Label2 & "','" & Format(DTPicker1.Value, "yyyy-MM-dd") & "', '" & Text2.Text & "', N'" & txtIssued.Text & "', 0,'" & CurrentDept & "','" & MyDDE.GetFieldByName("Order ID") & "','M1')"
   
       .PrepareUpdate = " UPDATE [backflush_Header]" & _
                        " Set Dept = N'" & lblSupplier(0) & "', DateTrans = CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), Note = N'" & Text2.Text & "', [Issued BY] = N'" & txtIssued.Text & "', [Received By] = N'" & txtApproved.Text & "', Status = " & mVarStatus & ",OrderID='" & MyDDE.GetFieldByName("Order ID") & "'" & _
                        " WHERE  (IDTrans = N'" & Label2 & "')"
   Else
       .PrepareAppend = " INSERT INTO [backflush_Header]" & _
                        " (IDTrans, DateTrans, Note, [Issued BY], Status,dept,typetrans)" & _
                        " VALUES  (N'" & Label2 & "','" & Format(DTPicker1.Value, "yyyy-MM-dd") & "','" & Text2.Text & "', N'" & txtIssued.Text & "', 0,'" & CurrentDept & "','M0')"
   
       .PrepareUpdate = " UPDATE [backflush_Header]" & _
                        " Set Dept = N'" & lblSupplier(0) & "', DateTrans = CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), Note = N'" & Text2.Text & "', [Issued BY] = N'" & txtIssued.Text & "', [Received By] = N'" & txtApproved.Text & "', Status = " & mVarStatus & _
                        " WHERE  (IDTrans = N'" & Label2 & "')"
   End If
                        
       .PrepareDelete = "DELETE FROM [backflush_Header] WHERE     (IDTrans = N'" & Label2 & "')"
End With
Err.Clear
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("IDTrans")), MyDDE.GetFieldByName("IDTrans"), "xxx")
If Not IsNull(MyDDE.GetFieldByName("dept")) Then
   lblSupplier(0).Caption = MyDDE.GetFieldByName("dept")
Else
   lblSupplier(0).Caption = " "
End If
txtApproved.Text = IIf(IsNull(MyDDE.GetFieldByName("approved_by")), "", MyDDE.GetFieldByName("approved_by"))
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
              ' If MyData.CheckGridKosong(MyDDE.ChildRecordset, "Saldo") = False And MyDDE.ChildRecordset.Recordcount <> 0 Then
                  MyDDE.IsChildMemberReady = True
                  'DGPurchase.Columns(7).Visible = True
                  'DGPurchase.Columns(6).Visible = False
              ' Else
              '    MyDDE.IsChildMemberReady = False
              '    MessageBox "Data detail belum ada/belum siap. Harap diisi dulu.", "Peringatan", msgOkOnly, msgCrtical
              ' End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
'            mVarAdd = False
       
'       Case tmbDelete:
'            If MyDDE.CheckEmptyControl = False Then
'               MyDDE.IsChildMemberReady = True
'            Else
'               MyDDE.IsChildMemberReady = False
'            End If
'            mVarAdd = False
'       Case tmbCancel: mVarAdd = False
       Case tmbDetail:
                If MyDDE.CheckEmptyControl = False Then
                   MyDDE.CancelTrans = False
                   MyDDE.IsChildMemberReady = True
'                   If MyDDE.ChildRecordset.Recordcount <> 0 Then mVarStage = IIf(IsNull(MyDDE.ChildRecordset.Fields("StageID")), "", MyDDE.ChildRecordset.Fields("StageID"))
'                   DGPurchase.Columns(7).Visible = True
 '                  DGPurchase.Columns(6).Visible = False
                Else
                   MyDDE.IsChildMemberReady = False
                   MyDDE.CancelTrans = True
                   'MessageBox "Tidak bisa menambah detail komponen.", "Peringatan", msgOkOnly
                End If
End Select
End Sub


Private Sub OpenDetail(ByVal ParameterString As String)
If ParameterString = "" Then ParameterString = "xxxxxxxx"
rs.DBOpen "select * from querySPB where IDTrans='" & ParameterString & "'", CNN, lckLockBatch
Set MyDDE.ChildRecordset = rs.DBRecordset
Set DGPurchase.DataSource = MyDDE.ChildRecordset
End Sub

Private Function OpenPartner(ByVal Index As Integer) As Boolean
Select Case Index
       Case 0: RcPartner.DBOpen "SELECT OrderID AS [Order ID], OrderName AS [Order Name], Type AS [Job Type], Status AS [Current Status] FROM [Manufacture Order] WHERE (Status = N'RELEASED')", CNN, lckLockReadOnly
       Case 1: RcPartner.DBOpen "SELECT noItem as Kode,InternalName as [Nama Barang],Merk,UOM as Satuan,NoGroup from inventory where Manufacture = 0 ", CNN, lckLockReadOnly
       Case 2: RcPartner.DBOpen "select kode_dept as Kode, dept as [Nama Dept] from dept", CNN, lckLockReadOnly
       Case 3: RcPartner.DBOpen "SELECT [backflush_Header].IDTrans AS [Order ID], [Manufacture Order].OrderName AS [Order Name], [Manufacture Order].Type AS [Job Type], [Manufacture Order].Status AS [Current Status] FROM         [Manufacture Order] INNER JOIN   [backflush_Header] ON [Manufacture Order].OrderID = [backflush_Header].OrderID", CNN, lckLockReadOnly
       Case 4: RcPartner.DBOpen "SELECT backflush_line.NoItem AS [Item ID], backflush_line.Description AS Description, backflush_line.UOM, backflush_line.Lokasi AS WareHouse, backflush_line.Cost AS Cost,  backflush_line.[Qty Required] AS [Quote Qty], backflush_line.StageID, backflush_line.ResourcesID FROM Inventory INNER JOIN backflush_line ON Inventory.NoItem = backflush_line.NoItem WHERE     (backflush_line.IDTrans = N'" & MyDDE.GetFieldByName(0) & "') ORDER BY backflush_line.NoItem", CNN, lckLockReadOnly
       Case 5: RcPartner.DBOpen "SELECT sl_no as [Batch / Lot No],StockTmp as [Sisa Stock] FROM [Inventory tabel] where noItem='" & MyDDE.ChildRecordset.Fields("noItem") & "' and StockTmp > 0 ", CNN, lckLockReadOnly
End Select
If RcPartner.Recordcount <> 0 Then
    Select Case Index
           Case 0: mCall.FromTagActive = "PRODUCTION ORDER"
           Case 1: mCall.FromTagActive = "Stage"
           Case 2: mCall.FromTagActive = "Departemen"
           Case 3: mCall.FromTagActive = "Manufacturing Order"
           Case 4: mCall.FromTagActive = "Detail Komponen"
           Case 5: mCall.FromTagActive = "Batch No / Lot No"
    End Select
    Set mCall.FormData = RcPartner.DBRecordset
    mCall.LookUp Me
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgCrtical
   If MyDDE.ChildRecordset.Recordcount <> 0 Then
       If IsNull(MyDDE.ChildRecordset.Fields("noItem")) = True Or MyDDE.ChildRecordset.Fields("noItem") = "" Then
          MyDDE.ChildRecordset.CancelBatch adAffectCurrent
          If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
       End If
    End If
   OpenPartner = True
End If
End Function

Private Function IndexAuto() As String
Dim Rc As New DBQuick
Dim TglSaiki As String
Dim Inom As Long
TglSaiki = Format(Day(Date), "0#") & Format(Month(Date), "0#") & Right(Format(Year(Date), "0#"), 2)
Rc.DBOpen "SELECT     MAX(RIGHT(IDTrans, 5)) AS MaxNom FROM         [backflush_Header] WHERE     (LEFT(IDTrans, 2) = N'RE')", CNN, lckLockReadOnly
With Rc
     If .DBRecordset.Recordcount <> 0 Then
        Inom = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + 1
     Else
        Inom = 1
     End If
     Select Case Len(Trim(Str(Inom)))
            Case 0: IndexAuto = "RE/" & TglSaiki & "-" & Trim(Str(Inom))
            Case 1: IndexAuto = "RE/" & TglSaiki & "-" & "0000" & Trim(Str(Inom))
            Case 2: IndexAuto = "RE/" & TglSaiki & "-" & "000" & Trim(Str(Inom))
            Case 3: IndexAuto = "RE/" & TglSaiki & "-" & "00" & Trim(Str(Inom))
            Case 4: IndexAuto = "RE/" & TglSaiki & "-" & "0" & Trim(Str(Inom))
     End Select
End With
End Function

Private Sub OpenKaryawan()
RcIssued.DBOpen "SELECT     EmpID AS [Issued BY], FullName AS [Nama Karyawan] FROM         Employees", CNN, lckLockReadOnly
'DataCombo1(0).ListField = "Nama Karyawan"
'Set DataCombo1(0).RowSource = RcIssued.DBRecordset

RcApprov.DBOpen "SELECT     EmpID AS [Received By], FullName AS [Nama Karyawan] FROM         Employees", CNN, lckLockReadOnly
'DataCombo1(1).ListField = "Nama Karyawan"
'Set DataCombo1(1).RowSource = RcApprov.DBRecordset
End Sub

Private Function KasihHarga(ByVal NoItem As String) As Long
Dim RcLok As New DBQuick
RcLok.DBOpen "SELECT     MIN(HPP) AS Harga FROM         [Inventory Tabel] WHERE     (NoItem = N'" & NoItem & "') AND (LEFT(RefTrans, 2) = N'RN')", CNN
With RcLok.DBRecordset
     If .Recordcount <> 0 Then
        KasihHarga = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        KasihHarga = 0
     End If
     RcLok.CloseDB
End With
Set RcLok = Nothing
End Function

Private Function QTYReceived(ByVal NoOrder As String, ByVal NoItem As String) As Long
Dim RcRec As New DBQuick
RcRec.DBOpen "SELECT     [Qty Received] FROM         backflush_line WHERE     (IDTRans = N'" & NoOrder & "') AND (NoItem = N'" & NoItem & "')", CNN, lckLockReadOnly
With RcRec.DBRecordset
     If .Recordcount <> 0 Then
         QTYReceived = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
         QTYReceived = 0
     End If
End With
End Function




