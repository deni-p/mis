VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmSFIssue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shop Floor Operation - Material Issue"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMatIssue.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   11445
   Tag             =   "Material Issue"
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   11
      Top             =   5130
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5220
      Left            =   0
      ScaleHeight     =   5220
      ScaleWidth      =   11505
      TabIndex        =   13
      Top             =   0
      Width           =   11505
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         DataField       =   "Received By"
         Height          =   330
         Left            =   1995
         MaxLength       =   249
         TabIndex        =   10
         Tag             =   "RN"
         Top             =   4395
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         DataField       =   "issued By"
         Height          =   330
         Left            =   1995
         MaxLength       =   249
         TabIndex        =   9
         Tag             =   "RN"
         Top             =   4035
         Width           =   2535
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   5370
         Picture         =   "FrmMatIssue.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   473
         Width           =   350
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         DataField       =   "Note"
         Height          =   330
         Left            =   1995
         MaxLength       =   249
         MultiLine       =   -1  'True
         TabIndex        =   8
         Tag             =   "RN"
         Top             =   3675
         Width           =   5250
      End
      Begin MSDataGridLib.DataGrid DGPurchase 
         Height          =   2340
         Left            =   105
         TabIndex        =   7
         Tag             =   "Partner"
         Top             =   1260
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   4128
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
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "Item ID"
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
            Caption         =   "Barang"
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
            DataField       =   "UOM"
            Caption         =   "UOM"
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
            DataField       =   "Lokasi"
            Caption         =   "Lokasi"
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
            Caption         =   "Qty Required"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#.##0,0000;(#.##0,0000)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Qty Warehouse"
            Caption         =   "Qty Warehouse"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#.##0,0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Qty Received"
            Caption         =   "Qty Received"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#.##0,0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "StageID"
            Caption         =   "StageID"
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
         BeginProperty Column08 
            DataField       =   "ResourcesID"
            Caption         =   "ResourcesID"
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
         BeginProperty Column09 
            DataField       =   "Cost"
            Caption         =   "Cost"
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
         BeginProperty Column10 
            DataField       =   "idx"
            Caption         =   "idx"
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   3825
         TabIndex        =   12
         Tag             =   "RN"
         Text            =   "Text1"
         Top             =   2895
         Visible         =   0   'False
         Width           =   885
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "DateTrans"
         Height          =   330
         Left            =   7590
         TabIndex        =   4
         Tag             =   "RN"
         Top             =   90
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
         Format          =   65077251
         CurrentDate     =   38538
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "IDTrans"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1995
         TabIndex        =   0
         Tag             =   "RN"
         Top             =   105
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
         Left            =   1995
         TabIndex        =   3
         Tag             =   "RN"
         Top             =   825
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
         Left            =   1995
         TabIndex        =   2
         Tag             =   "RN"
         Top             =   465
         Width           =   3375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         DataField       =   "Tanggal"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   6225
         TabIndex        =   22
         Top             =   158
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NoTrans"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   7
         Left            =   195
         TabIndex        =   21
         Top             =   173
         Width           =   600
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   150
         X2              =   2760
         Y1              =   3990
         Y2              =   3990
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   6210
         X2              =   8040
         Y1              =   405
         Y2              =   405
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   150
         X2              =   2490
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Received by"
         DataField       =   "Tanggal"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   6
         Left            =   195
         TabIndex        =   20
         Top             =   4463
         Width           =   885
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   150
         X2              =   2490
         Y1              =   4710
         Y2              =   4710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issued By"
         DataField       =   "Tanggal"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   5
         Left            =   195
         TabIndex        =   19
         Top             =   4103
         Width           =   705
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   150
         X2              =   2460
         Y1              =   4350
         Y2              =   4350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Catatan"
         DataField       =   "Tanggal"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   3
         Left            =   195
         TabIndex        =   18
         Top             =   3743
         Width           =   585
      End
      Begin VB.Label lblSupplier 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Job Type"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   3
         Left            =   7590
         TabIndex        =   5
         Tag             =   "RN"
         Top             =   465
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
         Left            =   7590
         TabIndex        =   6
         Tag             =   "RN"
         Top             =   825
         Width           =   3720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Type"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   4
         Left            =   6225
         TabIndex        =   17
         Top             =   533
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   16
         Top             =   893
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Manufacture Order"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   15
         Top             =   533
         Width           =   1380
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   150
         X2              =   2700
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   150
         X2              =   2730
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   6210
         X2              =   8040
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   6210
         X2              =   8115
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Status"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   10
         Left            =   6225
         TabIndex        =   14
         Top             =   893
         Width           =   1065
      End
   End
End
Attribute VB_Name = "FrmSFIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As New DBQuick
Private RcPartner As New DBQuick
Private RcIssued As New DBQuick
Private RcApprov As New DBQuick

Private WithEvents mCall As frmCallerLogistik
Attribute mCall.VB_VarHelpID = -1
Private WithEvents mCall1 As frmCaller
Attribute mCall1.VB_VarHelpID = -1


Public Property Let Mode(isLogistik As Boolean)
   If isLogistik Then
      lblSupplier(0).Visible = False
      lblSupplier(1).Visible = False
      lblSupplier(2).Visible = False
      lblSupplier(3).Visible = False
      CmdLink(0).Visible = False
      Label1(0).Visible = False
      Label1(1).Visible = False
      Label1(4).Visible = False
      Label1(10).Visible = False
      Line1(0).Visible = False
      Line1(1).Visible = False
      Line1(4).Visible = False
      Line1(5).Visible = False
   Else
      lblSupplier(0).Visible = True
      lblSupplier(1).Visible = True
      lblSupplier(2).Visible = True
      lblSupplier(3).Visible = True
      CmdLink(0).Visible = True
      Label1(0).Visible = True
      Label1(1).Visible = True
      Label1(4).Visible = True
      Label1(10).Visible = True
      Line1(0).Visible = True
      Line1(1).Visible = True
      Line1(4).Visible = True
      Line1(5).Visible = True
   End If
End Property

Private Sub GridLayout()
With DGPurchase
    'APPEARANCE
    .Columns(6).Visible = False
    .Columns(7).Visible = False
    .Columns(8).Visible = False
    .Columns(9).Visible = False
    
    'WIDTH
    .Columns(0).width = 1500    'NoItem
    .Columns(1).width = 3000
    .Columns(2).width = 1000
    .Columns(3).width = 1870
    .Columns(4).width = 1600
    .Columns(5).width = 1600
    'CAPTION
    .Columns(4).Caption = "Qty Request"
    .Columns(5).Caption = "Qty Issued"
    'ALIGNMENT
    .Columns(4).Alignment = dbgRight
    .Columns(5).Alignment = dbgRight
    'FORMAT
    .Columns(4).NumberFormat = QtyFormFloat
    .Columns(5).NumberFormat = QtyFormFloat
End With
'DGPurchase.Columns(5).Width = 1514.835
End Sub

Private Sub cmdLink_Click(Index As Integer)
OpenPartner Index
End Sub

Private Sub DGPurchase_AfterColUpdate(ByVal ColIndex As Integer)
Select Case ColIndex
       Case 5:
            If DGPurchase.Columns(ColIndex) = "" Then DGPurchase.Columns(ColIndex) = 0
            If Val(DGPurchase.Columns(5)) > Val(DGPurchase.Columns(4)) Then DGPurchase.Columns(5) = DGPurchase.Columns(4)
'            If CDbl(DGPurchase.Columns(ColIndex).Value) > (MyDDE.ChildRecordset.Fields("Qty Received") - CekStockPO(MyDDE.ChildRecordset.Fields("Item ID"))) Then
'               MessageBox "Quantity tidak boleh lebih besar dari nilai kuota yang dibutuhkan bagian produksi", "Peringatan", msgOkOnly, msgCrtical
'               DGPurchase.Columns(ColIndex).Value = 0
'            End If
End Select

End Sub

Private Sub DGPurchase_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'If DataCombo1(0).Enabled = True Then
On Error GoTo xErr
Select Case DGPurchase.col
       Case 5:
            DGPurchase.MarqueeStyle = dbgFloatingEditor
            DGPurchase.AllowUpdate = True
            
       Case Else
            DGPurchase.MarqueeStyle = dbgFloatingEditor
            DGPurchase.AllowUpdate = False
End Select
Exit Sub
xErr:
   Err.Clear
'End If
End Sub

Private Sub Form_Load()
GridLayout
HiasFormManTell Picture2, Me
'OpenKaryawan
With MyDDE
    .EditModeReplace = False
    Set .BindForm = Me
    .SetPermissions = UserDeleteDenied
    .BindFormTAG = "RN"
    Set .ActiveConnection = CNN
    
    '.PrepareQuery = " SELECT     [backflush_Header].OrderID AS [Order ID], [Manufacture Order].OrderName AS [Nama Order], [Manufacture Order].Type AS [Job Type], [Manufacture Order].Status AS [Current Status], [backflush_Header].DateTrans, [backflush_Header].Note, [backflush_Header].[Issued BY], [backflush_Header].[Received By], [backflush_Header].IDTrans" & _
                    " FROM         [Manufacture Order] INNER JOIN  [backflush_Header] ON [Manufacture Order].OrderID = [backflush_Header].OrderID WHERE     ([backflush_Header].Status = 0) AND  (LEFT([backflush_Header].IDTrans, 2) = N'MR') ORDER BY [backflush_Header].OrderID"
     
     .PrepareQuery = " SELECT [backflush_Header].OrderID AS [Order ID], [Manufacture Order].OrderName AS [Nama Order], " & _
                             "[Manufacture Order].Type AS [Job Type], [Manufacture Order].Status AS [Current Status], " & _
                             "[backflush_Header].DateTrans, [backflush_Header].Note, [backflush_Header].[Issued BY], " & _
                             "[backflush_Header].[Received By], [backflush_Header].IDTrans,[backflush_Header].[Doc Ref]" & _
                     " FROM [Manufacture Order] Right outer JOIN " & _
                           "[backflush_Header] ON [Manufacture Order].OrderID = [backflush_Header].OrderID " & _
                     " WHERE ([backflush_Header].Status = 0) AND (LEFT([backflush_Header].IDTrans, 2) = N'MR') " & _
                     " ORDER BY [backflush_Header].OrderID "
   

End With
Set mCall = New frmCallerLogistik
Set mCall1 = New frmCaller
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
Set FrmSFIssue = Nothing
End Sub

Private Sub mCall_BeforeUnload()
Select Case mCall.FromTagActive
       Case "DETAIL MATERIAL ISSUED", "DETAIL FACTORY REQUEST":
            If FindOwnRecordset(MyDDE.ChildRecordset, "[Item ID] = '" & MyDDE.ChildRecordset.Fields("Item ID") & "'") = True Then
               MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Item ID") & " Sudah Ada....! Silahkan Diulangi", "Peringatan", msgOkOnly, msgCrtical
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
Select Case UCase(TagForm)
       Case "MANUFACTURING ORDER":
            With MyDDE
                 .GetFieldByName(0) = mCall.GetFieldByName(0)
                 .GetFieldByName(1) = mCall.GetFieldByName(1)
                 .GetFieldByName(2) = mCall.GetFieldByName(2)
                 .GetFieldByName(3) = mCall.GetFieldByName(3)
            End With
       Case "DETAIL MATERIAL ISSUED", "DETAIL FACTORY REQUEST":
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
                        MyDDE.ChildRecordset.Fields("item ID") = .GetFieldByName("item ID")
                        MyDDE.ChildRecordset.Fields("description") = .GetFieldByName("IDTRans")
                        MyDDE.ChildRecordset.Fields("UOM") = .GetFieldByName("uom")
                        MyDDE.ChildRecordset.Fields("Lokasi") = .GetFieldByName("lokasi")
                        MyDDE.ChildRecordset.Fields("Qty WareHouse") = 0
                        MyDDE.ChildRecordset.Fields("Qty Required") = .GetFieldByName("quote Qty")
                        MyDDE.ChildRecordset.Fields("Qty Received") = .GetFieldByName("quote Qty")
                        MyDDE.ChildRecordset.Fields("StageID") = .GetFieldByName("StageID")
                        MyDDE.ChildRecordset.Fields("ResourcesID") = .GetFieldByName("ResourcesID")
                        MyDDE.ChildRecordset.Fields("Cost") = .GetFieldByName("Cost")
                        MyDDE.ChildRecordset.Fields("idx") = .GetFieldByName("idx")
                        MyDDE.ChildRecordset.Fields("internalName") = .GetFieldByName("InternalName")
                        'messagebox .GetFieldByName("Quote qty")
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
       Case tmbAddNew:
            Dim IDGen As New IDGenerator
            Label2 = IDGen.GetID("MR")
            Set IDGen = Nothing
            DTPicker1.Value = Now
            MyDDE.GetFieldByName("Note") = "-"
            MyDDE.GetFieldByName("issued By") = MainMenu.StatusBar1.Panels(1).Text
            MyDDE.GetFieldByName("Received By") = "-"
       Case tmbDetail:
            If MyDDE.IsChildMemberReady = True Then
               If lblSupplier(0).Caption = "" Then
                  OpenPartner 2
               Else
                  OpenPartner 1
               End If
            End If
       Case tmbDelete:
       
            If MyDDE.IsChildMemberReady = True Then

            End If
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               Dim aStatus As String
               With MyDDE.ChildRecordset
               If .Recordcount <> 0 Then
                   If SendDataToServer("DELETE FROM  [backflush_line] WHERE     (IDTrans = N'" & Label2 & "')") = True Then
                        .MoveFirst
                        Do
                          If .EOF Then Exit Do
                            ' SendDataToServer (" UPDATE    backflush_line" & _
                                               " Set [Qty Warehouse] = " & CDbl(.Fields("Qty Warehouse")) & _
                                               " WHERE     (IDTrans = N'" & Label2 & "') AND (NoItem = N'" & .Fields("Item ID") & "')")
                            aStatus = IIf(DGPurchase.Columns(4) = DGPurchase.Columns(5), "1", "0")
                            
                            SendDataToServer "update backflush_line set [Qty Required]=" & Val(DGPurchase.Columns(4)) - Val(DGPurchase.Columns(5)) & ",status=" & aStatus & " where idx='" & .Fields("idx") & "'"
                            
                            SendDataToServer (" INSERT INTO backflush_line (IDTrans, StageID,                                                          OrderID,                    NoItem,                       Description,                       UOM,                        Lokasi,                       Cost,                          [Qty Warehouse],                          [Qty Received],                      [Qty Required],                          ResourcesID)" & _
                                                        " VALUES (N'" & Label2 & "', N'" & IIf(IsNull(.Fields("StageID")), "", .Fields("StageID")) & "', N'" & lblSupplier(0) & "', N'" & .Fields("Item ID") & "', N'" & .Fields("Description") & "', N'" & .Fields("UOM") & "', N'" & .Fields("Lokasi") & "',  " & CDbl(.Fields("Cost")) & ",  " & CDbl(.Fields("Qty Warehouse")) & "," & CDbl(.Fields("Qty Received")) & "," & .Fields("Qty Received") & ", N'" & .Fields("ResourcesID") & "')")
                            
                            SendDataToServer "update backflush_header set status=1 where IDTrans='" & .Fields("Description") & "'"
                            
                            If lblSupplier(0).Caption = "" Then
                              UpdateStock .Fields("Item ID"), Val(DGPurchase.Columns(5))
                            Else
                              SendARItem .Fields("Item ID"), CDbl(.Fields("Qty Warehouse")), CDbl(.Fields("Cost")), Label2, DTPicker1.Value, CDbl(.Fields("Cost")), "MR", True
                              SendQTY .Fields("Item ID")
                            End If
                          .MoveNext
                        Loop
                        .MoveFirst
                        ClosedPO
                   End If
               End If
               End With
            End If
       Case tmbPrint:
            CallRPTReport "Job Issues Report.Rpt", "Select * from [Detail Job Issues Report] where [IDTRANS]=N'" & Label2 & "'", , , "Select * from [Detail Job Issues Report]", "Detail Job Issues Report"
       Case tmbEdit:
            
End Select
CmdLink(0).Enabled = Text1.Enabled
End Sub

Private Sub UpdateStock(ID As String, Value As Double)
   Dim rsItem As New DBQuick
   Dim sisa As Double
   rsItem.DBOpen "select Noidx,stocktmp from [inventory tabel] where noItem='" & ID & "' and LockFifo=0", CNN, lckLockBatch
   With rsItem.DBRecordset
      If .Recordcount > 0 Then
         .MoveFirst
         sisa = Value
         While Not .EOF
            If sisa > 0 Then
               If Val(.Fields("stockTmp")) <= sisa Then
                  SendDataToServer "update [inventory tabel] set stockTmp=0, lockFifo=1, Qty_Out=Qty_out +" & sisa & " where Noidx='" & .Fields("Noidx") & "'"
                  sisa = sisa - Val(.Fields("stockTmp"))
               Else
                  SendDataToServer "update [inventory tabel] set stockTmp=stockTmp -" & sisa & ", Qty_Out=Qty_out +" & sisa & " where Noidx='" & .Fields("Noidx") & "'"
                  sisa = 0
               End If
            End If
            .MoveNext
         Wend
      End If
   End With
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
With MyDDE
    .PrepareAppend = " INSERT INTO [backflush_Header]" & _
                     " (IDTrans, OrderID, DateTrans, Note, [Issued BY], [Received By], Status, TypeTrans)" & _
                     " VALUES  (N'" & Label2 & "', N'" & lblSupplier(0) & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), N'" & Text2.Text & "', N'" & Text3.Text & "', N'" & Text4.Text & "', 0,'MI')"
'MessageBox .PrepareAppend
    .PrepareUpdate = " UPDATE [backflush_Header]" & _
                     " Set OrderID = N'" & lblSupplier(0) & "', DateTrans = CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), Note = N'" & Text2.Text & "', [Issued BY] = N'" & Text3.Text & "', [Received By] = N'" & Text4.Text & "', Status = 0" & _
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


Private Sub OpenDetail(ByVal ParameterString As String)

If ParameterString = "" Then ParameterString = "xxxxxxxx"

rs.DBOpen "SELECT backflush_line.NoItem AS [Item ID], backflush_line.Description, backflush_line.UOM, backflush_line.Lokasi, backflush_line.[Qty Required], backflush_line.[Qty Warehouse], backflush_line.[Qty Received], backflush_line.StageID, backflush_line.ResourcesID, backflush_line.Cost, backflush_line.IDx, inventory.internalName " & _
    " From backflush_line inner join inventory on backflush_line.noItem=inventory.noItem WHERE (IDTrans = N'" & ParameterString & "') ORDER BY [Item ID]", CNN


'"SELECT NoItem AS [Item ID], Description, UOM, Lokasi, [Qty Warehouse] AS [Qty Warehouse], [Qty Received], StageID, ResourcesID,Cost FROM backflush_line WHERE     (IDTrans = N'" & ParameterString & "') ORDER BY NoItem", CNN, lckLockBatch

Set MyDDE.ChildRecordset = rs.DBRecordset   '.Clone(adLockBatchOptimistic)
Set DGPurchase.DataSource = MyDDE.ChildRecordset
End Sub

Private Function OpenPartner(ByVal Index As Integer) As Boolean
Select Case Index
       Case 0:
'           RcPartner.DBOpen "SELECT [backflush_Header].IDTrans AS [Order ID], [Manufacture Order].OrderName AS [Order Name]," & _
'                                    "[Manufacture Order].Type AS [Job Type],[Manufacture Order].Status AS [Current Status] " & _
'                             "FROM [Manufacture Order] INNER JOIN " & _
'                                  "[backflush_Header] ON [Manufacture Order].OrderID = [backflush_Header].OrderID", CNN, lckLockReadOnly
            RcPartner.DBOpen "SELECT [backflush_header].IDTrans AS [Order ID], [Manufacture Order].OrderName AS [Order Name], " & _
                                    "[Manufacture Order].Type AS [Job Type],[Manufacture Order].Status AS [Current Status] " & _
                             "FROM [Manufacture Order] INNER JOIN  " & _
                                  "[backflush_header] ON [Manufacture Order].OrderID = [backflush_header].OrderID", CNN, lckLockReadOnly
                                  
       Case 1:
           'RcPartner.DBOpen "SELECT backflush_line.NoItem AS [Item ID], backflush_line.Description AS Description, backflush_line.UOM, backflush_line.Lokasi AS WareHouse, backflush_line.Cost AS Cost,  backflush_line.[Qty Required] AS [Quote Qty], backflush_line.StageID, backflush_line.ResourcesID FROM         Inventory INNER JOIN backflush_line ON Inventory.NoItem = backflush_line.NoItem WHERE     (backflush_line.IDTrans = N'" & MyDDE.GetFieldByName(0) & "') ORDER BY backflush_line.NoItem", CNN, lckLockReadOnly
           RcPartner.DBOpen "SELECT  backflush_line.NoItem AS [Item ID], backflush_line.Description AS Description, backflush_line.[Qty Required] AS [Quote Qty], backflush_line.UOM, backflush_line.Lokasi AS WareHouse, backflush_line.Cost AS Cost,   backflush_line.StageID, backflush_line.ResourcesID,backflush_line.idx FROM         Inventory INNER JOIN backflush_line ON Inventory.NoItem = backflush_line.NoItem WHERE     (backflush_line.IDTrans = N'" & MyDDE.GetFieldByName(0) & "') ORDER BY backflush_line.NoItem", CNN, lckLockReadOnly
       Case 2:
           RcPartner.DBOpen "SELECT backflush_header.DateTrans AS Tanggal, Inventory.InternalName , " & _
                               " backflush_line.[Qty Required] AS [Quote Qty], backflush_line.UOM, backflush_header.[Issued BY], backflush_line.Cost, " & _
                               " backflush_line.StageID, backflush_line.ResourcesID, backflush_line.IDX, backflush_header.dept, " & _
                               " backflush_line.NoItem AS [Item ID], backflush_line.Lokasi AS WareHouse, backflush_line.penggunaan, backflush_line.description, backflush_line.IDTRans " & _
                           " FROM Inventory INNER JOIN " & _
                               " backflush_line ON Inventory.NoItem = backflush_line.NoItem INNER JOIN " & _
                               " backflush_header ON backflush_line.IDTrans = backflush_header.IDTrans " & _
                           " WHERE (LEFT(backflush_line.IDTrans, 2) = 'FR') AND (backflush_line.Status <> 1) AND (backflush_header.TypeTrans <> N'MI') AND (backflush_line.[Qty Required] <> 0) and backflush_header.approved_by is not null" & _
                           " ORDER BY backflush_header.DateTrans desc , backflush_header.[Issued BY]", CNN, lckLockReadOnly
                           
          '" WHERE (LEFT(backflush_line.IDTrans, 2) = 'FR') AND (backflush_line.Status <> 1) AND (backflush_header.TypeTrans <> N'MI') and (backflush_header.approved_by is not null)"
End Select

If RcPartner.Recordcount <> 0 Then
    Select Case Index
           Case 0: mCall1.FromTagActive = "Manufacturing Order"
           Case 1: mCall1.FromTagActive = "Detail Material Issued"
           Case 2: mCall.FromTagActive = "Detail Factory Request"
    End Select
    If Index = 2 Then
      Set mCall.FormData = RcPartner.DBRecordset
       mCall.LookUp Me
    Else
      Set mCall1.FormData = RcPartner.DBRecordset
      mCall1.LookUp Me
    End If
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgCrtical
   OpenPartner = True
End If
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

Private Function CekStockPO(ByVal NoItem As String) As Long
Dim Rc As New DBQuick
Rc.DBOpen "SELECT     SUM([Qty Warehouse]) AS [Qty Warehouse] FROM         backflush_line WHERE     (LEFT(IDTrans, 2) = N'MR') AND (NoItem = N'" & NoItem & "') AND (OrderID = N'" & lblSupplier(0) & "')", CNN, lckLockReadOnly
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
RcQtyTotal.DBOpen "SELECT     SUM([Qty Warehouse]) AS [Qty Warehouse] FROM         backflush_line WHERE     (LEFT(IDTrans, 2) = N'MR') AND (OrderID = N'" & lblSupplier(0) & "') AND (NoItem = N'" & NoItem & "')", CNN, lckLockReadOnly
With RcQtyTotal.DBRecordset
     If .Recordcount <> 0 Then
        SendDataToServer (" UPDATE    backflush_line" & _
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
RcCari.DBOpen "SELECT     OrderID FROM         [backflush_Header] WHERE     (IDTrans = N'" & lblSupplier(0) & "')", CNN, lckLockReadOnly
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
RcClose.DBOpen " SELECT     SUM([Qty Required]) - SUM([Qty Received]) AS Complete, OrderID FROM         backflush_line WHERE     (IDTrans = N'" & lblSupplier(0) & "') GROUP BY OrderID ", CNN, lckLockReadOnly
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


