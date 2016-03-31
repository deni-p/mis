VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmSFBackflush 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shop Floor Operation - Backflush WIP"
   ClientHeight    =   5355
   ClientLeft      =   225
   ClientTop       =   1455
   ClientWidth     =   11055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBackFlushOutput.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   Tag             =   "BackFlushing"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4785
      Left            =   0
      ScaleHeight     =   4785
      ScaleWidth      =   11055
      TabIndex        =   11
      Top             =   0
      Width           =   11055
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         DataField       =   "Received By"
         Height          =   330
         Left            =   5580
         MaxLength       =   249
         TabIndex        =   22
         Tag             =   "RN"
         Top             =   4155
         Width           =   2640
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         DataField       =   "Issued BY"
         Height          =   330
         Left            =   1455
         MaxLength       =   249
         TabIndex        =   21
         Tag             =   "RN"
         Top             =   4155
         Width           =   2640
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   4845
         Picture         =   "frmBackFlushOutput.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   458
         Width           =   330
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         DataField       =   "Note"
         Height          =   330
         Left            =   1455
         MaxLength       =   249
         TabIndex        =   8
         Tag             =   "RN"
         Top             =   3795
         Width           =   6765
      End
      Begin VB.TextBox Text1 
         DataField       =   "Order ID"
         Height          =   330
         Left            =   2370
         TabIndex        =   10
         Tag             =   "RN"
         Text            =   "Text1"
         Top             =   1995
         Visible         =   0   'False
         Width           =   645
      End
      Begin MSDataGridLib.DataGrid DGPurchase 
         Bindings        =   "frmBackFlushOutput.frx":6BDC
         Height          =   2415
         Left            =   75
         TabIndex        =   9
         Top             =   1275
         Width           =   10785
         _ExtentX        =   19024
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
            DataField       =   "StageID"
            Caption         =   "Stage"
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
            DataField       =   "NoItem"
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
         BeginProperty Column02 
            DataField       =   "Description"
            Caption         =   "Keterangan"
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
            DataField       =   "UOM"
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
         BeginProperty Column04 
            DataField       =   "Lokasi"
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
         BeginProperty Column05 
            DataField       =   "Quantity Input"
            Caption         =   "Qty Input"
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
            DataField       =   "Qty Received"
            Caption         =   "Qty Terima"
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
         SplitCount      =   1
         BeginProperty Split0 
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
            BeginProperty Column05 
               Alignment       =   1
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "DateTrans"
         Height          =   330
         Left            =   1455
         TabIndex        =   3
         Tag             =   "RN"
         Top             =   795
         Width           =   3720
         _ExtentX        =   6562
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
         Format          =   65863683
         CurrentDate     =   38538
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "IDTrans"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1455
         TabIndex        =   1
         Tag             =   "RN"
         Top             =   90
         Width           =   3720
      End
      Begin VB.Label lblSupplier 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Job Type"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   3
         Left            =   7110
         TabIndex        =   6
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
         Left            =   7110
         TabIndex        =   7
         Tag             =   "RN"
         Top             =   810
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
         Left            =   7110
         TabIndex        =   5
         Tag             =   "RN"
         Top             =   90
         Width           =   3720
      End
      Begin VB.Label lblSupplier 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Order ID"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   1455
         TabIndex        =   2
         Tag             =   "RN"
         Top             =   450
         Width           =   3390
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   90
         X2              =   1600
         Y1              =   405
         Y2              =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Trans"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   20
         Top             =   180
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approval By"
         DataField       =   "Tanggal"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   6
         Left            =   4290
         TabIndex        =   19
         Top             =   4230
         Width           =   870
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   4245
         X2              =   5695
         Y1              =   4470
         Y2              =   4470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issued By"
         DataField       =   "Tanggal"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   4230
         Width           =   705
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   90
         X2              =   1600
         Y1              =   4470
         Y2              =   4470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note"
         DataField       =   "Tanggal"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   3870
         Width           =   345
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   90
         X2              =   1600
         Y1              =   4110
         Y2              =   4110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         DataField       =   "Tanggal"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   870
         Width           =   570
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   90
         X2              =   1600
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Type"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   4
         Left            =   5805
         TabIndex        =   15
         Top             =   518
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Order"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   5805
         TabIndex        =   14
         Top             =   158
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MO. Number "
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   525
         Width           =   945
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   90
         X2              =   1600
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   7300
         X2              =   5760
         Y1              =   405
         Y2              =   405
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   7300
         X2              =   5760
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   7300
         X2              =   5760
         Y1              =   1125
         Y2              =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Status"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   10
         Left            =   5805
         TabIndex        =   12
         Top             =   885
         Width           =   1065
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   4785
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmSFBackflush"
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

Public Property Let SetTagForm(ByVal vNewValue As Variant)
       Me.Tag = ""
       Me.Tag = vNewValue
'       Picture1.Print Me.Tag
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
End Property

Private Sub cmdLink_Click(Index As Integer)
OpenPartner Index
End Sub

Private Sub DataGrid1_Click()

End Sub

Private Sub DGPurchase_AfterColEdit(ByVal ColIndex As Integer)
'If cmdLink(0).Enabled = True Then
'    Select Case ColIndex
'           Case 5:
''                If CDbl(DGPurchase.Columns(ColIndex).Value) > 0 Then
'                 DGPurchase.Columns(7).Value = Val(DGPurchase.Columns(7).Value) + Val(DGPurchase.Columns(5).Value)   'QTYReceived(Label2, MyDDE.ChildRecordset("Item ID")) + DGPurchase.Columns(ColIndex).Value
''                   DGPurchase.Columns(4).Value = 0
''                End If
'    End Select
'End If
End Sub

Private Sub DGPurchase_ButtonClick(ByVal ColIndex As Integer)
If cmdLink(0).Enabled = True And ColIndex = 4 Then If MyDDE.ActiveRecordset.Recordcount <> 0 Then OpenPartner 2
End Sub

Private Sub DGPurchase_Error(ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub

Private Sub DGPurchase_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If cmdLink(0).Enabled = True Then
    Select Case DGPurchase.col
           Case 4, 5:
                DGPurchase.AllowUpdate = True
           Case Else
                DGPurchase.AllowUpdate = False
    End Select
Else
    DGPurchase.AllowUpdate = False
End If
End Sub

Private Sub Form_Load()
GridLayout
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
'OpenKaryawan
With MyDDE
    .EditModeReplace = False
    Set .BindForm = Me
    .BindFormTAG = "RN"
    .SetPermissions = UserDeleteDenied
    Set .ActiveConnection = CNN
'    If Me.Tag = "MATERIAL ISSUED" Then
'       .PrepareQuery = "SELECT     [backflush_Header].OrderID AS [Order ID], [Manufacture Order].OrderName AS [Nama Order], [Manufacture Order].Type AS [Job Type],     [Manufacture Order].Status AS [Current Status], [backflush_Header].DateTrans, [backflush_Header].Note, [backflush_Header].[Issued BY],                        [backflush_Header].[Received By], [backflush_Header].IDTrans FROM         [Manufacture Order] INNER JOIN                       [backflush_Header] ON [Manufacture Order].OrderID = [backflush_Header].OrderID WHERE     ([backflush_Header].Status = 0) ORDER BY [backflush_Header].IDTrans, [backflush_Header].OrderID"
'    Else

'       .PrepareQuery = " SELECT  [backflush_Header].OrderID AS [Order ID], [Manufacture Order].OrderName AS [Nama Order], [Manufacture Order].Type AS [Job Type], [Manufacture Order].Status AS [Current Status], [backflush_Header].DateTrans, [backflush_Header].Note, [backflush_Header].[Issued BY],  [backflush_Header].[Received By], [backflush_Header].IDTrans FROM [backflush_Header] [backflush_Header_1] INNER JOIN" & _
'                       " [backflush_Header] ON [backflush_Header_1].IDTrans = [backflush_Header].OrderID INNER JOIN [Manufacture Order] ON [backflush Header_1].OrderID = [Manufacture Order].OrderID WHERE     ([backflush_Header].Status = 1) AND (LEFT([backflush_Header].IDTrans, 3) = N'BFL') ORDER BY [backflush_Header].IDTrans, [backflush_Header].OrderID"
       
       .PrepareQuery = " SELECT  [backflush_Header].OrderID AS [Order ID], [Manufacture Order].OrderName AS [Nama Order], " & _
                                "[Manufacture Order].Type AS [Job Type], [Manufacture Order].Status AS [Current Status], " & _
                                "[backflush_Header].DateTrans, [backflush_Header].Note, [backflush_Header].[Issued BY], " & _
                                "[backflush_Header].[Received By], [backflush_Header].IDTrans " & _
                       " FROM [backflush_Header] INNER JOIN " & _
                            " [Manufacture Order] ON [backflush_Header].OrderID = [Manufacture Order].OrderID " & _
                       " WHERE ([backflush_Header].Status = 1) AND (LEFT([backflush_Header].IDTrans, 3) = N'BFL') ORDER BY [backflush_Header].IDTrans, [backflush_Header].OrderID"


'    End If
End With
Set mCall = New frmCaller
lblSupplier(0).ForeColor = vbWindowText
lblSupplier(1).ForeColor = vbWindowText
lblSupplier(2).ForeColor = vbWindowText
lblSupplier(3).ForeColor = vbWindowText
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
Set frmSFBackflush = Nothing
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case TagForm
       Case "PRODUCTION ORDER":
            With MyDDE
                 .GetFieldByName(0) = mCall.GetFieldByName(0)
                 .GetFieldByName(1) = mCall.GetFieldByName(1)
                 .GetFieldByName(2) = mCall.GetFieldByName(2)
                 .GetFieldByName(3) = mCall.GetFieldByName(3)
            End With
                  Dim RcCompData As New DBQuick
                  Dim I As Integer
                  
                  RcCompData.DBOpen "SELECT backflush_line.StageID, backflush_line.NoItem, backflush_line.Description, " & _
                                " backflush_line.UOM, inventory.warehouse as Lokasi, backflush_line.[Quantity Input], backflush_line.[Qty Received] - backflush_line.[Quantity Input] AS [Qty Received]," & _
                                " backflush_line.[Qty Required],dbo.backflush_line.Cost,Inventory.NoAccount,Inventory.ItemNAme " & _
                                " FROM backflush_line INNER JOIN Inventory ON backflush_line.NoItem = Inventory.NoItem " & _
                                " WHERE (backflush_line.IDTrans = N'" & lblSupplier(0) & "') AND (Inventory.Manufacture = 0) AND (backflush_line.[Qty Received] - backflush_line.[Quantity Input] > 0)", CNN, lckLockReadOnly
                  Debug.Print RcCompData.DBRecordset.Source
                  With RcCompData.DBRecordset
                       If .Recordcount > 0 Then
                           I = 0
                           OpenDetail "xxx" 'IIf(Not IsNull(MyDDE.GetFieldByName("No Order")), MyDDE.GetFieldByName("No Order"), "xxx")
                           Do
                             I = I + 1
                             If .EOF Then Exit Do
                             MyDDE.ChildRecordset.AddNew
                             MyDDE.ChildRecordset.Fields("StageID") = .Fields(0)
                             MyDDE.ChildRecordset.Fields("NoItem") = .Fields(1)
                             MyDDE.ChildRecordset.Fields("Description") = .Fields("ItemName")
                             MyDDE.ChildRecordset.Fields("UOM") = .Fields(3)
                             MyDDE.ChildRecordset.Fields("lokasi") = .Fields(4)
                             MyDDE.ChildRecordset.Fields("Quantity Input") = 0
                             MyDDE.ChildRecordset.Fields("Qty Received") = .Fields("Qty Received")
                             MyDDE.ChildRecordset.Fields("Qty Required") = .Fields("Qty Required")
                             MyDDE.ChildRecordset.Fields("Cost") = .Fields("Cost")
                             MyDDE.ChildRecordset.Fields("NoAccount") = .Fields("NoAccount")
                             .MoveNext
                           Loop
                           MyDDE.ChildRecordset.MoveFirst
                       End If
                  End With
       Case "WAREHOUSE":
             MyDDE.ChildRecordset.Fields("Lokasi") = mCall.GetFieldByName(0)
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
                            MyDDE.ChildRecordset.Fields(4) = 0
                            MyDDE.ChildRecordset.Fields(5) = .GetFieldByName(4)
                            MyDDE.ChildRecordset.Fields(6) = .GetFieldByName(5)
                            MyDDE.ChildRecordset.Fields(7) = 0
                            MyDDE.ChildRecordset.Fields(8) = .GetFieldByName(6)
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
            MyDDE.GetFieldByName("Issued By") = MainMenu.StatusBar1.Panels(1).Text
            MyDDE.GetFieldByName("Received By") = "-"
            DTPicker1.Value = Date
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
                    If SendDataToServer("DELETE FROM  [backflush_line] WHERE     (IDTrans = N'" & Label2 & "')") = True Then
                         Dim sJournal As New clsJournal
                         Dim oIDJ As New IDGenerator
                         Dim sIDj As String
                         sIDj = oIDJ.GetID("MM")
                         sJournal.CiptaKaryaHeaderJournal sIDj, Label2.Caption, "", "", MainMenu.StatusBar1.Panels(1), "", "IDR", Now, mVarPeriode, "MEMORIAL"
                         Set oIDJ = Nothing
                         
                         .MoveFirst
                         Do
                           If .EOF Then Exit Do
                              SendDataToServer (" INSERT INTO  [backflush_line]" & _
                                                " (Idtrans,StageID,OrderID, NoItem, Description, UOM, Lokasi,   [Quantity Input], [Qty Received],[Qty Required],cost)" & _
                                                " VALUES  (N'" & Label2 & "',N'" & .Fields("StageID") & "',N'" & lblSupplier(0) & "', N'" & .Fields("NoItem") & "', N'" & .Fields("Description") & "', N'" & .Fields("UOM") & "', N'" & .Fields("Lokasi") & "', " & CDbl(.Fields("Quantity Input")) & "," & CDbl(.Fields("Qty Received")) & "," & CDbl(.Fields("Qty Required")) & "," & CDbl(.Fields("Cost")) & ")")
                              SendAPItem .Fields("noitem"), CDbl(.Fields("Quantity Input")), CDbl(0), Label2, DTPicker1.Value, "BFL", 0, 0, True, .Fields("lokasi")
                              sJournal.CiptaKaryaDetailJournal sIDj, .Fields("NoAccount"), "", Val(.Fields("cost")) * Val(.Fields("Quantity Input")), 0, "Barang Dalam Proses " & " " & .Fields("noItem") & " " & .Fields("Description")
                              sJournal.CiptaKaryaDetailJournal sIDj, .Fields("NoAccount"), "", 0, Val(.Fields("cost")) * Val(.Fields("Quantity Input")), .Fields("Description")
                           .MoveNext
                         Loop
                         .MoveFirst
                     End If
                End If
               End With
            End If
       Case tmbPrint:
            CallRPTReport "BackFlushing.Rpt", "Select * From [BackFlushing] where IDTRANS ='" & Label2 & "'"
End Select
cmdLink(0).Enabled = Text1.Enabled
DGPurchase.Columns(4).Button = cmdLink(0).Enabled
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
Dim mVarStatus As Byte
With MyDDE
    If Me.Tag = "MATERIAL ISSUED" Then
       mVarStatus = 0
    Else
       mVarStatus = 1
    End If
    .PrepareAppend = " INSERT INTO [backflush_Header]" & _
                     " (IDTrans, OrderID, DateTrans, Note, [Issued BY], [Received By], Status)" & _
                     " VALUES  (N'" & Label2 & "', N'" & lblSupplier(0) & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), N'" & Text2.Text & "', N'" & Text3.Text & "', N'" & Text4.Text & "', 1)"
'MessageBox .PrepareAppend
    .PrepareUpdate = " UPDATE [backflush_Header]" & _
                     " Set OrderID = N'" & lblSupplier(0) & "', DateTrans = CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), Note = N'" & Text2.Text & "', [Issued BY] = N'" & Text3.Text & "', [Received By] = N'" & Text4.Text & "', Status = 1" & _
                     " WHERE  (IDTrans = N'" & Label2 & "')"
                     
    .PrepareDelete = " DELETE FROM [backflush_Header] WHERE     (IDTrans = N'" & Label2 & "')"
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
               If MyDDE.ChildRecordset.Recordcount <> 0 Then
                  MyDDE.IsChildMemberReady = True
               Else
                  MyDDE.IsChildMemberReady = False
                  MessageBox "Data detail belum ada/belum siap. Harap diisi dulu.", "Peringatan", msgOkOnly, msgCrtical
               End If
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

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

Private Sub OpenDetail(ByVal ParameterString As String)
Dim rs As New DBQuick
If ParameterString = "" Then ParameterString = "xxxxxxxx"
'rs.DBOpen " SELECT BackFlush.NoItem AS [Item ID], BackFlush.Description, BackFlush.UOM, BackFlush.Lokasi, BackFlush.[Quantity Input], BackFlush.Cost,  BackFlush.[Qty Required], BackFlush.[Qty Received], BackFlush.StageID FROM         BackFlush INNER JOIN [Order Output Detail] ON BackFlush.StageID = [Order Output Detail].StageID" & _
          " WHERE (BackFlush.IDTrans = N'" & ParameterString & "') GROUP BY BackFlush.NoItem, BackFlush.Description, BackFlush.UOM, BackFlush.Lokasi, BackFlush.[Quantity Input], BackFlush.Cost,   BackFlush.[Qty Required], BackFlush.[Qty Received], BackFlush.StageID, [Order Output Detail].SeqNo ORDER BY [Order Output Detail].SeqNo", Cnn, lckLockBatch
          
rs.DBOpen "SELECT backflush_line.StageID, backflush_line.NoItem, backflush_line.Description, backflush_line.UOM, backflush_line.Lokasi, backflush_line.[Quantity Input], backflush_line.[Qty Received] AS [Qty Received], backflush_line.IDTrans,backflush_line.[Qty Required],backflush_line.cost,inventory.noAccount FROM  backflush_line inner join inventory on backflush_line.noItem = inventory.noitem WHERE (IDTrans = N'" & ParameterString & "')", CNN, lckLockBatch

Set MyDDE.ChildRecordset = rs.DBRecordset.Clone(adLockBatchOptimistic)
Set DGPurchase.DataSource = MyDDE.ChildRecordset
End Sub

Private Function OpenPartner(ByVal Index As Integer) As Boolean
Select Case Index
       Case 0:
         'RcPartner.DBOpen "SELECT     [backflush_Header].IDTrans, [backflush_Header].OrderID, [Manufacture Order].OrderName, [Manufacture Order].Type, [Manufacture Order].Status FROM         [backflush_Header] INNER JOIN [Manufacture Order] ON [backflush_Header].OrderID = [Manufacture Order].OrderID WHERE     (LEFT([backflush_Header].IDTrans, 2) = N'RE') AND ([Manufacture Order].Status = N'RELEASED')", CNN, lckLockReadOnly
         RcPartner.DBOpen "SELECT [backflush_header].IDTrans, [backflush_header].OrderID, [Manufacture Order].OrderName, " & _
                                 "[Manufacture Order].Type, [Manufacture Order].Status " & _
                          "FROM [backflush_header] INNER JOIN " & _
                               "[Manufacture Order] ON [backflush_header].OrderID = [Manufacture Order].OrderID " & _
                          "WHERE (LEFT([backflush_header].IDTrans, 2) = N'RE') AND ([Manufacture Order].Status = N'RELEASED')", CNN, lckLockReadOnly

       Case 1:
         'RcPartner.DBOpen "SELECT     [Ord Comp Detail].NoItem AS [Item ID], [Ord Comp Detail].[DESC] AS Description, [Ord Comp Detail].UOM, Inventory.WareHouse AS WareHouse,                       0 AS Cost, [Ord Comp Detail].[Quote Qty],[Ord Comp Detail].[StageID] FROM         [Ord Comp Detail] INNER JOIN                       Inventory ON [Ord Comp Detail].NoItem = Inventory.NoItem WHERE     ([Ord Comp Detail].OrderID = N'" & MyDDE.GetFieldByName(0) & "') order by  [Ord Comp Detail].NoItem", CNN, lckLockReadOnly
         RcPartner.DBOpen "SELECT [Ord Comp Detail].NoItem AS [Item ID], [Ord Comp Detail].[DESC] AS Description, " & _
                                 "[Ord Comp Detail].UOM, Inventory.WareHouse AS WareHouse, 0 AS Cost, " & _
                                 "[Ord Comp Detail].[Quote Qty],[Ord Comp Detail].[StageID] " & _
                          "FROM [Ord Comp Detail] INNER JOIN Inventory ON [Ord Comp Detail].NoItem = Inventory.NoItem " & _
                          "WHERE ([Ord Comp Detail].OrderID = N'" & MyDDE.GetFieldByName(0) & "') " & _
                          "ORDER BY  [Ord Comp Detail].NoItem", CNN, lckLockReadOnly
       
       
       Case 2:
         'RcPartner.DBOpen "SELECT     WareHouse, [WareHouse Name] FROM         WareHouse ORDER BY WareHouse", CNN, lckLockReadOnly
         RcPartner.DBOpen "SELECT WareHouse, [WareHouse Name] FROM WareHouse ORDER BY WareHouse", CNN, lckLockReadOnly

End Select
If RcPartner.Recordcount <> 0 Then
    Select Case Index
           Case 0: mCall.FromTagActive = "PRODUCTION ORDER"
           Case 1: mCall.FromTagActive = "COMPONENT DETAIL"
           Case 2: mCall.FromTagActive = "WAREHOUSE"
    End Select
    Set mCall.FormData = RcPartner.DBRecordset
    mCall.LookUp Me
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgCrtical
   OpenPartner = True
End If
End Function

Private Function IndexAuto() As String
Dim Rc As New DBQuick
Dim TglSaiki As String
Dim Inom As Long
TglSaiki = Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2)
Rc.DBOpen "SELECT     MAX(RIGHT(IDTrans, 5)) AS MaxNom FROM         [backflush_Header] WHERE     (LEFT(IDTrans, 3) = N'BFL')", CNN, lckLockReadOnly
With Rc
     If .DBRecordset.Recordcount <> 0 Then
        Inom = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + 1
     Else
        Inom = 1
     End If
     Select Case Len(Trim(Str(Inom)))
            Case 0: IndexAuto = "BFL/" & TglSaiki & "-" & Trim(Str(Inom))
            Case 1: IndexAuto = "BFL/" & TglSaiki & "-" & "0000" & Trim(Str(Inom))
            Case 2: IndexAuto = "BFL/" & TglSaiki & "-" & "000" & Trim(Str(Inom))
            Case 3: IndexAuto = "BFL/" & TglSaiki & "-" & "00" & Trim(Str(Inom))
            Case 4: IndexAuto = "BFL/" & TglSaiki & "-" & "0" & Trim(Str(Inom))
     End Select
End With
End Function

'Private Sub OpenKaryawan()
'RcIssued.DBOpen "SELECT     EmpID AS [Issued BY], FullName AS [Nama Karyawan] FROM         Employees", CNN, lckLockReadOnly
'DataCombo1(0).ListField = "Nama Karyawan"
'Set DataCombo1(0).RowSource = RcIssued.DBRecordset
'
'RcApprov.DBOpen "SELECT     EmpID AS [Received By], FullName AS [Nama Karyawan] FROM         Employees", CNN, lckLockReadOnly
'DataCombo1(1).ListField = "Nama Karyawan"
'Set DataCombo1(1).RowSource = RcApprov.DBRecordset
'End Sub

Private Function LokasiTrans(ByVal NoStage As String) As String
Dim RcLok As New DBQuick
RcLok.DBOpen "SELECT WareHouse FROM         [Order Output Detail] WHERE     (OrderID = N'" & lblSupplier(0) & "') AND (StageID = N'" & NoStage & "') GROUP BY WareHouse", CNN
With RcLok.DBRecordset
     If .Recordcount <> 0 Then
        LokasiTrans = IIf(Not IsNull(.Fields(0)), .Fields(0), "")
     Else
        LokasiTrans = "-"
     End If
End With
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


Private Function LoadQty(ByVal NoItem As String) As Long
Dim Rcl As New DBQuick
Rcl.DBOpen "SELECT      SUM([Qty Received]) AS [Qty Received] FROM         backflush_line WHERE     (OrderID = N'" & Label2 & "') AND (NoItem = N'" & NoItem & "')  GROUP BY [Qty Received]", CNN, lckLockReadOnly
With Rcl.DBRecordset
     If .Recordcount <> 0 Then
        LoadQty = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        LoadQty = 0
     End If
End With
Rcl.CloseDB
Set Rcl = Nothing
End Function

Private Sub GridLayout()
DGPurchase.width = 10785
DGPurchase.Height = 2415
DGPurchase.Columns(0).width = 1425.26
DGPurchase.Columns(1).width = 1755.213
DGPurchase.Columns(2).width = 2415.118
DGPurchase.Columns(3).width = 675.2126
DGPurchase.Columns(4).width = 1649.764
DGPurchase.Columns(5).width = 1140.095
DGPurchase.Columns(6).width = 1140.095
'DGPurchase.Columns(7).Width = 1110.047
End Sub

