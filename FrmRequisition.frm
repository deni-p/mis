VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{594F23A7-88F5-4C02-866B-8E877A62F75C}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmRequisition 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   13140
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   13140
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5355
      Left            =   90
      ScaleHeight     =   5325
      ScaleWidth      =   12960
      TabIndex        =   12
      Top             =   0
      Width           =   12990
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   4560
         Left            =   -45
         ScaleHeight     =   4530
         ScaleWidth      =   12660
         TabIndex        =   13
         Top             =   255
         Width           =   12690
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            DataField       =   "Note"
            Height          =   330
            Left            =   1875
            MaxLength       =   249
            TabIndex        =   8
            Tag             =   "RN"
            Top             =   3660
            Width           =   6765
         End
         Begin VB.CommandButton cmdLink 
            Enabled         =   0   'False
            Height          =   330
            Index           =   0
            Left            =   6000
            Picture         =   "FrmRequisition.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   450
            Width           =   405
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "DateTrans"
            Height          =   330
            Left            =   7905
            TabIndex        =   4
            Tag             =   "RN"
            Top             =   105
            Width           =   3720
            _ExtentX        =   6562
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MMMM/yyyy"
            Format          =   48758787
            CurrentDate     =   38538
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            DataField       =   "Issued BY"
            Height          =   330
            Index           =   0
            Left            =   1875
            TabIndex        =   9
            Tag             =   "RN"
            Top             =   4005
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Nama Karyawan"
            BoundColumn     =   "Issued BY"
            Text            =   "DataCombo1"
         End
         Begin MSDataGridLib.DataGrid DGPurchase 
            Height          =   2415
            Left            =   105
            TabIndex        =   7
            Top             =   1215
            Width           =   12420
            _ExtentX        =   21908
            _ExtentY        =   4260
            _Version        =   393216
            AllowUpdate     =   0   'False
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
            ColumnCount     =   6
            BeginProperty Column00 
               DataField       =   "Item ID"
               Caption         =   "Item ID"
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
               DataField       =   "Description"
               Caption         =   "Description"
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
               DataField       =   "UOM"
               Caption         =   "UOM"
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
            BeginProperty Column03 
               DataField       =   "Lokasi"
               Caption         =   "Lokasi"
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
            BeginProperty Column04 
               DataField       =   "Qty Warehouse"
               Caption         =   "QTY"
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
            BeginProperty Column05 
               DataField       =   "Qty Received"
               Caption         =   "Required"
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
                  ColumnWidth     =   2174.74
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   3929.953
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1335.118
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   2564.788
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnWidth     =   1800
               EndProperty
               BeginProperty Column05 
                  Object.Visible         =   0   'False
               EndProperty
            EndProperty
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            DataField       =   "Received By"
            Height          =   330
            Index           =   1
            Left            =   5985
            TabIndex        =   10
            Tag             =   "RN"
            Top             =   4005
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Nama Karyawan"
            BoundColumn     =   "Received By"
            Text            =   "DataCombo1"
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Left            =   3825
            TabIndex        =   14
            Tag             =   "RN"
            Text            =   "Text1"
            Top             =   2895
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   6540
            X2              =   7920
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
            Left            =   2685
            TabIndex        =   0
            Tag             =   "RN"
            Top             =   105
            Width           =   3720
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   510
            X2              =   2850
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NoTrans"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   7
            Left            =   510
            TabIndex        =   23
            Top             =   165
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Received by"
            DataField       =   "Tanggal"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   6
            Left            =   4635
            TabIndex        =   22
            Top             =   4065
            Width           =   990
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   4635
            X2              =   6000
            Y1              =   4320
            Y2              =   4320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Issued By"
            DataField       =   "Tanggal"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   5
            Left            =   510
            TabIndex        =   21
            Top             =   4065
            Width           =   780
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   510
            X2              =   1875
            Y1              =   4320
            Y2              =   4320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Note"
            DataField       =   "Tanggal"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   3
            Left            =   510
            TabIndex        =   20
            Top             =   3735
            Width           =   405
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   510
            X2              =   1875
            Y1              =   3975
            Y2              =   3975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal"
            DataField       =   "Tanggal"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   2
            Left            =   6540
            TabIndex        =   19
            Top             =   150
            Width           =   645
         End
         Begin VB.Label lblSupplier 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Job Type"
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   3
            Left            =   7905
            TabIndex        =   5
            Tag             =   "RN"
            Top             =   450
            Width           =   3720
         End
         Begin VB.Label lblSupplier 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Current Status"
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   2
            Left            =   7905
            TabIndex        =   6
            Tag             =   "RN"
            Top             =   795
            Width           =   3720
         End
         Begin VB.Label lblSupplier 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Nama Order"
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   1
            Left            =   2685
            TabIndex        =   3
            Tag             =   "RN"
            Top             =   795
            Width           =   3720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Job Type"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   4
            Left            =   6540
            TabIndex        =   18
            Top             =   510
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   1
            Left            =   510
            TabIndex        =   17
            Top             =   855
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Production Order Number "
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   0
            Left            =   510
            TabIndex        =   16
            Top             =   495
            Width           =   2175
         End
         Begin VB.Label lblSupplier 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Order ID"
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   0
            Left            =   2685
            TabIndex        =   1
            Tag             =   "RN"
            Top             =   450
            Width           =   3285
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   510
            X2              =   3060
            Y1              =   765
            Y2              =   765
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   510
            X2              =   3090
            Y1              =   1110
            Y2              =   1110
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   6540
            X2              =   8370
            Y1              =   765
            Y2              =   765
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   6540
            X2              =   8445
            Y1              =   1110
            Y2              =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Status"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   10
            Left            =   6540
            TabIndex        =   15
            Top             =   855
            Width           =   1200
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   11
      Top             =   5835
      Width           =   13140
      _ExtentX        =   23178
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmRequisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As New DBQuick
Private RcPartner As New DBQuick
Private RcIssued As New DBQuick
Private RcApprov As New DBQuick

Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1

Private Sub cmdLink_Click(Index As Integer)
OpenPartner Index
End Sub

Private Sub DGPurchase_AfterColEdit(ByVal ColIndex As Integer)
Select Case ColIndex
       Case 4:
            If CDbl(DGPurchase.Columns(ColIndex).Value) > (MyDDE.ChildRecordset.Fields("Qty Received") - CekStockPO(MyDDE.ChildRecordset.Fields("Item ID"))) Then
               MessageBox "Quantity tidak boleh lebih besar dari nilai kuota yang dibutuhkan bagian produksi", "Peringatan", msgOkOnly
               DGPurchase.Columns(ColIndex).Value = 0
            End If
End Select
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If DataCombo1(0).Enabled = True Then
Select Case DGPurchase.Col
       Case 4:
            DGPurchase.MarqueeStyle = dbgFloatingEditor
            DGPurchase.AllowUpdate = True
       Case Else
            DGPurchase.MarqueeStyle = dbgFloatingEditor
            DGPurchase.AllowUpdate = False
End Select
End If
End Sub

Private Sub Form_Load()
HiasForm Picture4, Me
CenterForm Picture2, Me
OpenKaryawan
With MyDDE

    .EditModeReplace = False
    Set .BindForm = FrmRequisition
    .SetPermissions = UserDeleteDenied
    .BindFormTAG = "RN"
    Set .ActiveConnection = Cnn
    '.PrepareQuery = " SELECT     [BackFlush Header].OrderID AS [Order ID], [Manufacture Order].OrderName AS [Nama Order], [Manufacture Order].Type AS [Job Type], [Manufacture Order].Status AS [Current Status], [BackFlush Header].DateTrans, [BackFlush Header].Note, [BackFlush Header].[Issued BY], [BackFlush Header].[Received By], [BackFlush Header].IDTrans" & _
                    " FROM         [Manufacture Order] INNER JOIN  [BackFlush Header] ON [Manufacture Order].OrderID = [BackFlush Header].OrderID WHERE     ([BackFlush Header].Status = 0) AND  (LEFT([BackFlush Header].IDTrans, 2) = N'MR') ORDER BY [BackFlush Header].OrderID"
     .PrepareQuery = " SELECT [BackFlush Header].OrderID AS [Order ID], [Manufacture Order].OrderName AS [Nama Order], [Manufacture Order].Type AS [Job Type], [Manufacture Order].Status AS [Current Status], [BackFlush Header].DateTrans, [BackFlush Header].Note, [BackFlush Header].[Issued BY], [BackFlush Header].[Received By], [BackFlush Header].IDTrans,[BackFlush Header].[Doc Ref]" & _
                     " FROM [BackFlush Header] [BackFlush Header_1] INNER JOIN [BackFlush Header] ON [BackFlush Header_1].IDTrans = [BackFlush Header].OrderID INNER JOIN  [Manufacture Order] ON [BackFlush Header_1].OrderID = [Manufacture Order].OrderID WHERE     ([BackFlush Header].Status = 0) AND (LEFT([BackFlush Header].IDTrans, 2) = N'MR') ORDER BY [BackFlush Header].OrderID"
End With
Set mCall = New frmCaller
lblSupplier(0).ForeColor = vbWindowText
lblSupplier(1).ForeColor = vbWindowText
lblSupplier(2).ForeColor = vbWindowText
lblSupplier(3).ForeColor = vbWindowText
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set rs = Nothing
Set RcPartner = Nothing
Set mCall = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmRequisition = Nothing
End Sub

Private Sub mCall_BeforeUnload()
Select Case mCall.FromTagActive
       Case "COMPONENT DETAIL":
            If FindOwnRecordset(MyDDE.ChildRecordset, "[Item ID] = '" & MyDDE.ChildRecordset.Fields("Item ID") & "'") = True Then
               MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Item ID") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
               MyDDE.ChildRecordset.CancelBatch adAffectCurrent
               If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
            Else
               If IsNull(MyDDE.ChildRecordset.Fields(0)) = True Or MyDDE.ChildRecordset.Fields(0) = "" Then
                  MyDDE.ChildRecordset.CancelBatch adAffectCurrent
                  If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
               End If
            End If
            If DGPurchase.Enabled = True Then DGPurchase.SetFocus
End Select
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case TagForm
       Case "MATERIAL ISSUED":
            With MyDDE
                 .GetFieldByName(0) = mCall.GetFieldByName(0)
                 .GetFieldByName(1) = mCall.GetFieldByName(1)
                 .GetFieldByName(2) = mCall.GetFieldByName(2)
                 .GetFieldByName(3) = mCall.GetFieldByName(3)
            End With
       Case "COMPONENT DETAIL":
'            Dim Rc As New DBQuick
'            Rc.DBOpen "SELECT     [Ord Comp Detail].NoItem AS [Item ID], [Ord Comp Detail].[DESC] AS Description, [Ord Comp Detail].UOM, Inventory.WareHouse AS WareHouse,                       0 AS Cost, [Ord Comp Detail].[Quote Qty],[Ord Comp Detail].[StageID] FROM         [Ord Comp Detail] INNER JOIN                       Inventory ON [Ord Comp Detail].NoItem = Inventory.NoItem WHERE     ([Ord Comp Detail].OrderID = N'" & MyDDE.GetFieldByName(0) & "') order by  [Ord Comp Detail].NoItem", Cnn
'            With Rc.DBRecordset
'                 If .Recordcount <> 0 Then
'                    With MyDDE.ChildRecordset
'                         If .Recordcount <> 0 Then
'                         .MoveFirst
'                          Do
'                            If .EOF Then Exit Do
'                            .Delete adAffectCurrent
'                            .MoveNext
'                          Loop
'                          End If
'                    End With
'                    Do
'                      If .EOF Then Exit Do
'                        MyDDE.ChildRecordset.AddNew
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
                        'MsgBox .GetFieldByName("Quote qty")
                        End With
'                       .MoveNext
'                    Loop
'                    MyDDE.ChildRecordset.MoveFirst
'                 End If
'            End With
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit:

       Case tmbAddNew:
            Label2 = IndexAuto
            MyDDE.GetFieldByName("Note") = "-"
       Case tmbDetail:
            If MyDDE.IsChildMemberReady = True Then
               OpenPartner 1
            End If
       Case tmbDelete:
       
            If MyDDE.IsChildMemberReady = True Then

            End If
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               With MyDDE.ChildRecordset
               If .Recordcount <> 0 Then
                   If SendDataToServer("DELETE FROM  [BackFlush] WHERE     (IDTrans = N'" & Label2 & "')") = True Then
                        .MoveFirst
                        Do
                          If .EOF Then Exit Do
                            ' SendDataToServer (" UPDATE    BackFlush" & _
                                               " Set [Qty Warehouse] = " & CDbl(.Fields("Qty Warehouse")) & _
                                               " WHERE     (IDTrans = N'" & Label2 & "') AND (NoItem = N'" & .Fields("Item ID") & "')")
                            SendDataToServer (" INSERT INTO BackFlush (IDTrans, StageID, OrderID, NoItem, Description, UOM, Lokasi,  Cost,  [Qty Warehouse],[Qty Received],[Qty Required],ResourcesID)" & _
                                              " VALUES (N'" & Label2 & "', N'" & .Fields("StageID") & "', N'" & lblSupplier(0) & "', N'" & .Fields("Item ID") & "', N'" & .Fields("Description") & "', N'" & .Fields("UOM") & "', N'" & .Fields("Lokasi") & "',  " & CDbl(.Fields("Cost")) & ",  " & CDbl(.Fields("Qty Warehouse")) & "," & CDbl(.Fields("Qty Received")) & "," & .Fields("Qty Received") & ", N'" & .Fields("ResourcesID") & "')")
                            SendQTY .Fields("Item ID")
                          .MoveNext
                        Loop
                        .MoveFirst
                        ClosedPO
                   End If
               End If
               End With
            End If
       Case tmbPrint:
            CallRPTReport "Material Requisition.Rpt", "Select * from [Material Requisition] where [IdTrans]=N'" & Label2 & "'"
End Select
cmdLink(0).Enabled = Text1.Enabled
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
With MyDDE
    .PrepareAppend = " INSERT INTO [BackFlush Header]" & _
                     " (IDTrans, OrderID, DateTrans, Note, [Issued BY], [Received By], Status)" & _
                     " VALUES  (N'" & Label2 & "', N'" & lblSupplier(0) & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), N'" & Text2 & "', N'" & DataCombo1(0).BoundText & "', N'" & DataCombo1(1).BoundText & "', 0)"
'MessageBox .PrepareAppend
    .PrepareUpdate = " UPDATE [BackFlush Header]" & _
                     " Set OrderID = N'" & lblSupplier(0) & "', DateTrans = CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), Note = N'" & Text2 & "', [Issued BY] = N'" & DataCombo1(0).BoundText & "', [Received By] = N'" & DataCombo1(1).BoundText & "', Status = 0" & _
                     " WHERE  (IDTrans = N'" & Label2 & "')"
    .PrepareDelete = " DELETE FROM [BackFlush Header] WHERE     (IDTrans = N'" & Label2 & "')"
End With
Err.Clear
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("IDTrans")), MyDDE.GetFieldByName("IDTrans"), "xxx")
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
'               If CekGrid = True And MyDDE.ChildRecordset.Recordcount <> 0 Then
                  MyDDE.IsChildMemberReady = True
'               Else
'                  MyDDE.IsChildMemberReady = False
'                  MessageBox "Data detail belum ada/belum siap. Harap diisi dulu.", "Peringatan", msgOkOnly
'               End If
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
                Else
                   MyDDE.IsChildMemberReady = False
                   MyDDE.CancelTrans = True
                   'MessageBox "Tidak bisa menambah detail komponen.", "Peringatan", msgOkOnly
                End If
End Select
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Picture4.Parent.hwnd
End Sub

Private Sub OpenDetail(ByVal ParameterString As String)
Dim rs As New DBQuick
If ParameterString = "" Then ParameterString = "xxxxxxxx"
rs.DBOpen " SELECT     NoItem AS [Item ID], Description, UOM, Lokasi, [Qty Warehouse] AS [Qty Warehouse], [Qty Received], StageID, ResourcesID,Cost FROM         BackFlush WHERE     (IDTrans = N'" & ParameterString & "') ORDER BY NoItem", Cnn, lckLockBatch
'MessageBox rs.DBRecordset.Source
Set MyDDE.ChildRecordset = rs.DBRecordset.Clone(adLockBatchOptimistic)
Set DGPurchase.DataSource = MyDDE.ChildRecordset
End Sub

Private Function OpenPartner(ByVal Index As Integer) As Boolean
Select Case Index
       Case 0:
            'RcPartner.DBOpen "SELECT     [BackFlush Header].OrderID AS [Order ID], [BackFlush Header].IDTrans AS [No Ref], [Manufacture Order].OrderName AS [Order Name], [Manufacture Order].Type AS [Job Type], [Manufacture Order].Status AS [Current Status] FROM         [Manufacture Order] INNER JOIN [BackFlush Header] ON [Manufacture Order].OrderID = [BackFlush Header].OrderID", Cnn, lckLockReadOnly
            RcPartner.DBOpen "SELECT     [BackFlush Header].IDTrans AS [Order ID], [Manufacture Order].OrderName AS [Order Name], [Manufacture Order].Type AS [Job Type],                       [Manufacture Order].Status AS [Current Status] FROM         [Manufacture Order] INNER JOIN   [BackFlush Header] ON [Manufacture Order].OrderID = [BackFlush Header].OrderID", Cnn, lckLockReadOnly
       Case 1: RcPartner.DBOpen "SELECT     BackFlush.NoItem AS [Item ID], BackFlush.Description AS Description, BackFlush.UOM, BackFlush.Lokasi AS WareHouse, BackFlush.Cost AS Cost,  BackFlush.[Qty Required] AS [Quote Qty], BackFlush.StageID, BackFlush.ResourcesID FROM         Inventory INNER JOIN BackFlush ON Inventory.NoItem = BackFlush.NoItem WHERE     (BackFlush.IDTrans = N'" & MyDDE.GetFieldByName(0) & "') ORDER BY BackFlush.NoItem", Cnn, lckLockReadOnly
End Select
If RcPartner.Recordcount <> 0 Then
    Select Case Index
           Case 0: mCall.FromTagActive = "MATERIAL ISSUED"
           Case 1: mCall.FromTagActive = "COMPONENT DETAIL"
    End Select
    Set mCall.FormData = RcPartner.DBRecordset
    mCall.LookUp Me
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly
   OpenPartner = True
End If
End Function

Private Function IndexAuto() As String
Dim Rc As New DBQuick
Dim TglSaiki As String
Dim Inom As Long
TglSaiki = Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2)
Rc.DBOpen "SELECT  MAX(RIGHT(IDTrans, 5)) AS MaxNom FROM         [BackFlush Header] [Manufacture Order] WHERE     (LEFT(IDTrans, 2) = N'MR')", Cnn, lckLockReadOnly
With Rc
     If .DBRecordset.Recordcount <> 0 Then
        Inom = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + 1
     Else
        Inom = 1
     End If
     Select Case Len(Trim(Str(Inom)))
            Case 0: IndexAuto = "MR/" & TglSaiki & "-" & Trim(Str(Inom))
            Case 1: IndexAuto = "MR/" & TglSaiki & "-" & "0000" & Trim(Str(Inom))
            Case 2: IndexAuto = "MR/" & TglSaiki & "-" & "000" & Trim(Str(Inom))
            Case 3: IndexAuto = "MR/" & TglSaiki & "-" & "00" & Trim(Str(Inom))
            Case 4: IndexAuto = "MR/" & TglSaiki & "-" & "0" & Trim(Str(Inom))
     End Select
End With
End Function

Private Sub OpenKaryawan()
RcIssued.DBOpen "SELECT     EmpID AS [Issued BY], FullName AS [Nama Karyawan] FROM         Employees", Cnn, lckLockReadOnly
DataCombo1(0).ListField = "Nama Karyawan"
Set DataCombo1(0).RowSource = RcIssued.DBRecordset

RcApprov.DBOpen "SELECT     EmpID AS [Received By], FullName AS [Nama Karyawan] FROM         Employees", Cnn, lckLockReadOnly
DataCombo1(1).ListField = "Nama Karyawan"
Set DataCombo1(1).RowSource = RcApprov.DBRecordset
End Sub

Private Function CekStockPO(ByVal NoItem As String) As Long
Dim Rc As New DBQuick
Rc.DBOpen "SELECT     SUM([Qty Warehouse]) AS [Qty Warehouse] FROM         BackFlush WHERE     (LEFT(IDTrans, 2) = N'MR') AND (NoItem = N'" & NoItem & "') AND (OrderID = N'" & lblSupplier(0) & "')", Cnn, lckLockReadOnly
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        CekStockPO = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        CekStockPO = 0
     End If
End With
Rc.CloseDB
Set Rc = Nothing
End Function

Private Sub SendQTY(ByVal NoItem As String)
Dim RcQtyTotal As New DBQuick
RcQtyTotal.DBOpen "SELECT     SUM([Qty Warehouse]) AS [Qty Warehouse] FROM         BackFlush WHERE     (LEFT(IDTrans, 2) = N'MR') AND (OrderID = N'" & lblSupplier(0) & "') AND (NoItem = N'" & NoItem & "')", Cnn, lckLockReadOnly
With RcQtyTotal.DBRecordset
     If .Recordcount <> 0 Then
        SendDataToServer (" UPDATE    BackFlush" & _
                          " Set [Qty Received] = " & CDbl(IIf(Not IsNull(.Fields(0)), .Fields(0), 0)) & _
                          " WHERE (IDTrans = N'" & lblSupplier(0) & "') AND (NoItem = N'" & NoItem & "')")
                          
        SendDataToServer (" UPDATE    [Ord Comp Detail]" & _
                          " Set [Actual Qty] = " & CDbl(IIf(Not IsNull(.Fields(0)), .Fields(0), 0)) & _
                          " WHERE  (OrderID = N'" & CariKode & "') AND (NoItem = N'" & NoItem & "')")
     End If
End With
RcQtyTotal.CloseDB
Set RcQtyTotal = Nothing
End Sub

Private Function CariKode() As String
Dim RcCari As New DBQuick
RcCari.DBOpen "SELECT     OrderID FROM         [BackFlush Header] WHERE     (IDTrans = N'" & lblSupplier(0) & "')", Cnn, lckLockReadOnly
With RcCari.DBRecordset
     If .Recordcount <> 0 Then
        CariKode = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        CariKode = "xxx"
     End If
     RcCari.CloseDB
End With
Set RcCari = Nothing
End Function

Private Function ClosedPO() As Boolean
Dim RcClose As New DBQuick
Dim ICek As Long
RcClose.DBOpen " SELECT     SUM([Qty Required]) - SUM([Qty Received]) AS Complete, OrderID FROM         BackFlush WHERE     (IDTrans = N'" & lblSupplier(0) & "') GROUP BY OrderID ", Cnn, lckLockReadOnly
With RcClose.DBRecordset
     If .Recordcount <> 0 Then
        If Not IsNull(.Fields(0)) Then
           If .Fields(0) = 0 Then
              SendDataToServer ("UPDATE [Manufacture Order] SET              Status = N'FINISHED' WHERE (OrderID = N'" & IIf(Not IsNull(.Fields(1)), .Fields(1), "XXXX") & "')")
           End If
        End If
     Else
     End If
     RcClose.CloseDB
End With
Set RcClose = Nothing
End Function


