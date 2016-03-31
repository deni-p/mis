VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmCustFeedBack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Feed Back"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCustFeedBack.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   9240
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   26
      Top             =   4605
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   1005
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
      Height          =   4665
      Left            =   0
      ScaleHeight     =   4665
      ScaleWidth      =   9240
      TabIndex        =   27
      Top             =   0
      Width           =   9240
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "approved_by"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   0
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   50
         Tag             =   "Feed"
         Top             =   4080
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Revisi"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   20
         Left            =   6030
         TabIndex        =   24
         Tag             =   "Feed"
         Top             =   3690
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Purchase Qty"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   9
         Left            =   1680
         TabIndex        =   13
         Tag             =   "Feed"
         Top             =   3345
         Width           =   735
      End
      Begin VB.CommandButton command1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   3135
         MaskColor       =   &H000000C0&
         Picture         =   "FrmCustFeedBack.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "minta_sampel"
         Top             =   1973
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.ComboBox ComboStatus 
         Appearance      =   0  'Flat
         DataField       =   "status complain"
         Height          =   315
         ItemData        =   "FrmCustFeedBack.frx":6BDC
         Left            =   6030
         List            =   "FrmCustFeedBack.frx":6BE6
         TabIndex        =   16
         Tag             =   "Feed"
         Top             =   934
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Complain Qty"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   10
         Left            =   6030
         TabIndex        =   14
         Tag             =   "Feed"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Batch No"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   11
         Left            =   6030
         TabIndex        =   15
         Tag             =   "Feed"
         Top             =   583
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Status Complain Qty"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   13
         Left            =   6030
         TabIndex        =   17
         Tag             =   "Feed"
         Top             =   1269
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "refNote"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   14
         Left            =   6030
         TabIndex        =   18
         Tag             =   "Feed"
         Top             =   1620
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Emp ID"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   15
         Left            =   6030
         TabIndex        =   19
         Tag             =   "Feed"
         Top             =   1965
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Receipt By"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   16
         Left            =   6030
         TabIndex        =   20
         Tag             =   "Feed"
         Top             =   2310
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Evaluation No"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   17
         Left            =   6030
         TabIndex        =   22
         Tag             =   "Feed"
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Doc No"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   19
         Left            =   6030
         TabIndex        =   23
         Tag             =   "Feed"
         Top             =   3345
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "Date Receipt"
         Height          =   300
         Index           =   0
         Left            =   6030
         TabIndex        =   21
         Tag             =   "Feed"
         Top             =   2655
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   57999363
         CurrentDate     =   39335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "Effective Date"
         Height          =   300
         Index           =   1
         Left            =   6030
         TabIndex        =   25
         Tag             =   "Feed"
         Top             =   4035
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   57999363
         CurrentDate     =   39335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Kepada"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   1
         Left            =   1680
         TabIndex        =   1
         Tag             =   "Feed"
         Top             =   583
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "tembusan"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   2
         Left            =   1680
         TabIndex        =   2
         Tag             =   "Feed"
         Top             =   926
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "lampiran"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   3
         Left            =   1680
         TabIndex        =   3
         Tag             =   "Feed"
         Top             =   1269
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "CompanyName"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   4
         Left            =   1680
         TabIndex        =   4
         Tag             =   "Feed"
         Top             =   1620
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "no item"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   5
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   6
         Tag             =   "Feed"
         Top             =   1965
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "kemasan"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   6
         Left            =   1680
         TabIndex        =   8
         Tag             =   "Feed"
         Top             =   2310
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "purchase ID"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   7
         Left            =   1680
         TabIndex        =   9
         Tag             =   "Feed"
         Top             =   2655
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Do id"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   8
         Left            =   1680
         TabIndex        =   11
         Tag             =   "Feed"
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "feedback id"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   18
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   0
         Tag             =   "Feed"
         Top             =   240
         Width           =   2310
      End
      Begin VB.CommandButton command1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   4215
         MaskColor       =   &H000000C0&
         Picture         =   "FrmCustFeedBack.frx":6BF7
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "minta_sampel"
         Top             =   1628
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton command1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   3375
         MaskColor       =   &H000000C0&
         Picture         =   "FrmCustFeedBack.frx":6F81
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "minta_sampel"
         Top             =   2663
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton command1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   3375
         MaskColor       =   &H000000C0&
         Picture         =   "FrmCustFeedBack.frx":730B
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "minta_sampel"
         Top             =   3008
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approval By"
         Height          =   195
         Index           =   22
         Left            =   180
         TabIndex        =   51
         Top             =   4140
         Width           =   870
      End
      Begin VB.Line Line1 
         Index           =   21
         X1              =   4680
         X2              =   6615
         Y1              =   4005
         Y2              =   4005
      End
      Begin VB.Line Line1 
         Index           =   20
         X1              =   4680
         X2              =   6615
         Y1              =   3660
         Y2              =   3660
      End
      Begin VB.Line Line1 
         Index           =   19
         X1              =   4680
         X2              =   7320
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line1 
         Index           =   18
         X1              =   4680
         X2              =   6615
         Y1              =   3315
         Y2              =   3315
      End
      Begin VB.Line Line1 
         Index           =   17
         X1              =   4680
         X2              =   6615
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Line Line1 
         Index           =   16
         X1              =   4680
         X2              =   6615
         Y1              =   2625
         Y2              =   2625
      End
      Begin VB.Line Line1 
         Index           =   15
         X1              =   4680
         X2              =   6615
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         Index           =   14
         X1              =   4680
         X2              =   6615
         Y1              =   1584
         Y2              =   1584
      End
      Begin VB.Line Line1 
         Index           =   13
         X1              =   4680
         X2              =   6615
         Y1              =   1234
         Y2              =   1234
      End
      Begin VB.Line Line1 
         Index           =   12
         X1              =   4680
         X2              =   6615
         Y1              =   898
         Y2              =   898
      End
      Begin VB.Line Line1 
         Index           =   11
         X1              =   4680
         X2              =   6615
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   4680
         X2              =   6615
         Y1              =   1935
         Y2              =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase QTY"
         Height          =   195
         Index           =   9
         Left            =   180
         TabIndex        =   49
         Top             =   3420
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   48
         Top             =   315
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kepada"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   47
         Top             =   645
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tembusan"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   46
         Top             =   990
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lampiran"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   45
         Top             =   1335
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   44
         Top             =   1695
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Item"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   43
         Top             =   2040
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kemasan"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   42
         Top             =   2385
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase ID"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   41
         Top             =   2730
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery  Order ID"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   40
         Top             =   3075
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Batch No"
         Height          =   195
         Index           =   11
         Left            =   4710
         TabIndex        =   38
         Top             =   645
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Complain"
         Height          =   195
         Index           =   12
         Left            =   4710
         TabIndex        =   37
         Top             =   990
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Complain QTY"
         Height          =   195
         Index           =   13
         Left            =   4710
         TabIndex        =   36
         Top             =   1335
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref Note"
         Height          =   195
         Index           =   14
         Left            =   4710
         TabIndex        =   35
         Top             =   1695
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Receipt"
         Height          =   195
         Index           =   15
         Left            =   4710
         TabIndex        =   34
         Top             =   2708
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Evaluation No"
         Height          =   195
         Index           =   16
         Left            =   4710
         TabIndex        =   33
         Top             =   3075
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Document"
         Height          =   195
         Index           =   17
         Left            =   4710
         TabIndex        =   32
         Top             =   3420
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Revisi"
         Height          =   195
         Index           =   18
         Left            =   4710
         TabIndex        =   31
         Top             =   3758
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date"
         Height          =   195
         Index           =   19
         Left            =   4710
         TabIndex        =   30
         Top             =   4088
         Width           =   1035
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   150
         X2              =   2085
         Y1              =   1935
         Y2              =   1935
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   150
         X2              =   2085
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   150
         X2              =   2085
         Y1              =   898
         Y2              =   898
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   150
         X2              =   2085
         Y1              =   1241
         Y2              =   1241
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   165
         X2              =   2100
         Y1              =   1584
         Y2              =   1584
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   150
         X2              =   2085
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   150
         X2              =   2085
         Y1              =   2625
         Y2              =   2625
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   150
         X2              =   2085
         Y1              =   2970
         Y2              =   2970
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   150
         X2              =   2085
         Y1              =   3315
         Y2              =   3315
      End
      Begin VB.Line Line1 
         Index           =   9
         X1              =   150
         X2              =   2085
         Y1              =   3660
         Y2              =   3660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emp ID"
         Height          =   195
         Index           =   20
         Left            =   4710
         TabIndex        =   29
         Top             =   2040
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt By"
         Height          =   195
         Index           =   21
         Left            =   4710
         TabIndex        =   28
         Top             =   2385
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Complain QTY"
         Height          =   195
         Index           =   10
         Left            =   4710
         TabIndex        =   39
         Top             =   308
         Width           =   990
      End
   End
End
Attribute VB_Name = "FrmCustFeedBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RcPartner As New DBQuick
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Dim sql As String

Private Sub ComboStatus_Change()
'If ComboStatus.Text = "-1" Then
'   ComboStatus.Text = "TRUE"
'Else: ComboStatus.Text = "0"
'    ComboStatus.Text = "FALSE"
'End If
End Sub

Private Sub Command1_Click(Index As Integer)
OpenPartner Index
End Sub

Private Sub mCall_CallLinkForm()
On Error GoTo 1
Select Case mCall.FromTagActive
       Case "MASTER CUSTOMER":
            'frmPartner.SetFocus
            'frmPartner.ZOrder (0)
       Case "PURCHASING":
            'frmBankPartner.SetFocus
            'frmBankPartner.ZOrder (0)
       Case "GUDANG":
            'frmWareHouse.SetFocus
            'frmWareHouse.ZOrder (0)
      ' Case "TERM METHOD"
End Select
Exit Sub
1:

End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
On Error GoTo 1
If pRecordset.Recordcount <> 0 Then
Select Case UCase(TagForm):
       Case "MASTER CUSTOMER":
            MyDDE.GetFieldByName("partner id") = mCall.GetFieldByName(0)
            Text1(4) = mCall.GetFieldByName(1)
       Case "PURCHASING"
            Text1(7) = mCall.GetFieldByName(0)
       Case "DO"
            Text1(8) = mCall.GetFieldByName(0)
       Case "MASTER BARANG"
             Text1(5) = mCall.GetFieldByName(0)
End Select
End If
Exit Sub
1:
MessageBox Err.Description, "frmcustfeedback_mcall_rowcolchange" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
Select Case AdReasonActiveDb
Case tmbAddNew
    Command1(0).Enabled = True
    Command1(1).Enabled = True
    Command1(2).Enabled = True
    Command1(3).Enabled = True
    DTPicker1(0).Enabled = True
    DTPicker1(1).Enabled = True
    Text1(18).Text = IndexAuto
Case tmbEdit
    Command1(0).Enabled = True
    Command1(1).Enabled = True
    Command1(2).Enabled = True
    Command1(3).Enabled = True
    DTPicker1(0).Enabled = True
    DTPicker1(1).Enabled = True
    Set DTPicker1(0).DataSource = MyDDE.ActiveRecordset
    Set DTPicker1(1).DataSource = MyDDE.ActiveRecordset
Case tmbCancel
    Command1(0).Enabled = False
    Command1(1).Enabled = False
    Command1(2).Enabled = False
    Command1(3).Enabled = False
    DTPicker1(0).Enabled = False
    DTPicker1(1).Enabled = False
 Case tmbPrint:
            Dim aReport As New utility
            aReport.CallReportView "select * from customerfeedback where [feedback id]='" & Text1(18) & "'", "customer feedback.rpt", ReportPath, "Customer FeedBack"
            Set aReport = Nothing
End Select
Exit Sub
Err.Clear
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
Select Case AdReasonActiveDb
Case tmbAddNew:
   Set DTPicker1(0).DataSource = MyDDE.ActiveRecordset
   Set DTPicker1(1).DataSource = MyDDE.ActiveRecordset
Case tmbSave
   MyDDE.IsChildMemberReady = True
   simpan
   DTPicker1(0).Enabled = False
   DTPicker1(1).Enabled = False
   Command1(0).Enabled = False
   Command1(1).Enabled = False
   Command1(2).Enabled = False
   Command1(3).Enabled = False
Case tmbDelete
   MyDDE.PrepareDelete = "delete from [Customer Feedback] where [feedback id] = '" & Text1(18).Text & "'"
   DTPicker1(0).Enabled = False
   DTPicker1(1).Enabled = False
Case tmbCancel
   DTPicker1(0).Enabled = False
   DTPicker1(1).Enabled = False
   Command1(0).Enabled = False
   Command1(1).Enabled = False
   Command1(2).Enabled = False
   Command1(3).Enabled = False
End Select
Exit Sub

End Sub
Function simpan()
On Error GoTo xErr
   MyDDE.PrepareAppend = "insert into [Customer Feedback] ([feedback ID], kepada, tembusan,lampiran, [partner id], [no item], kemasan, [purchase ID], [Do ID], [purchase Qty], [Complain Qty], [Batch No],[Status Complain],[Status Complain Qty], refNote,[Emp ID], [Receipt By], [Date Receipt], [Evaluation No], [Doc No], Revisi, [Effective Date],ordered_by) values " & _
                       " ('" & Text1(18).Text & "', '" & Text1(1).Text & "', '" & Text1(2).Text & "','" & Text1(3).Text & "',N'" & MyDDE.GetFieldByName("partner id") & "','" & Text1(5).Text & "', '" & Text1(6).Text & "', '" & Text1(7).Text & "', '" & Text1(8).Text & "', '" & Text1(9).Text & "', '" & Text1(10).Text & "', '" & Text1(11).Text & "','" & ComboStatus.Text & "', '" & Text1(13).Text & "', " & _
                       " '" & Text1(14).Text & "', '" & Text1(15).Text & "',  '" & Text1(16).Text & "', '" & Format(DTPicker1(0).Value, "yyyy-MM-dd") & "', '" & Text1(17).Text & "','" & Text1(19).Text & "','" & Text1(20).Text & "', '" & Format(DTPicker1(1).Value, "yyyy-MM-dd") & "','" & MainMenu.StatusBar1.Panels(1).Text & "')"
                       
                       
 MyDDE.PrepareUpdate = "update [Customer feedback] set kepada = '" & Text1(1).Text & "', tembusan = '" & Text1(2).Text & "', lampiran = '" & Text1(3).Text & "', [partner id] = N'" & MyDDE.GetFieldByName("partner id") & "', " & _
                     " [no item] = '" & Text1(5).Text & "', kemasan = '" & Text1(6).Text & "', [purchase ID] = '" & Text1(7).Text & "', " & _
                     " [Do id] = '" & Text1(8).Text & "', [purchase Qty] = '" & Text1(9).Text & "', [complain Qty] = '" & Text1(10).Text & "', [Batch No] = '" & Text1(11).Text & "', [status Complain] = '" & ComboStatus.Text & "', " & _
                     " [Status Complain Qty] = '" & Text1(13).Text & "', refNote = '" & Text1(14).Text & "', [Emp ID] = '" & Text1(15).Text & "', [Receipt By] = '" & Text1(16).Text & "', [Date Receipt] = '" & Format(DTPicker1(0).Value, "yyyy-MM-dd") & "' , [Evaluation No] =  '" & Text1(17).Text & "' , " & _
                     " [Doc No] = '" & Text1(19).Text & "' , Revisi = '" & Text1(20).Text & "', [Effective Date] = '" & Format(DTPicker1(1).Value, "yyyy-MM-dd") & "', ordered_by='" & MainMenu.StatusBar1.Panels(1).Text & "'  where [feedback ID] = '" & Text1(18).Text & "'"
Exit Function
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Function

Private Sub Form_Load()
On Error Resume Next
HiasFormManTell Picture2, Me
DTPicker1(0).Value = Now
DTPicker1(1).Value = Now

sql = "SELECT     dbo.[Customer Feedback].[Feedback ID], dbo.[Customer Feedback].Kepada, dbo.[Customer Feedback].tembusan, dbo.[Customer Feedback].Lampiran," & _
                      " dbo.[Customer Feedback].[partner ID], dbo.PartnerDB.CompanyName, dbo.PartnerDB.ContactName, dbo.[Customer Feedback].[No Item]," & _
                      " dbo.[Customer Feedback].Kemasan, dbo.[Customer Feedback].[Purchase ID], dbo.[Customer Feedback].[Do ID]," & _
                      " dbo.[Customer Feedback].[Purchase Qty], dbo.[Customer Feedback].[Complain Qty], dbo.[Customer Feedback].[Batch No]," & _
                      " dbo.[Customer Feedback].[Status Complain], dbo.[Customer Feedback].[Status Complain Qty], dbo.[Customer Feedback].refNote," & _
                      " dbo.[Customer Feedback].[Emp ID], dbo.[Customer Feedback].[Receipt By], dbo.[Customer Feedback].[Date Receipt]," & _
                      " dbo.[Customer Feedback].[Evaluation No], dbo.[Customer Feedback].[Doc No], dbo.[Customer Feedback].Revisi," & _
                      " dbo.[Customer Feedback].[Effective Date],dbo.[Customer Feedback].approved_by" & _
                      " FROM   dbo.[Customer Feedback] INNER JOIN" & _
                      " dbo.PartnerDB ON dbo.[Customer Feedback].[partner ID] = dbo.PartnerDB.PartnerID "

Set mCall = New frmCaller
Set MyDDE.ActiveConnection = CNN
Set MyDDE.BindForm = Me
    MyDDE.BindFormTAG = "Feed"
    MyDDE.PrepareQuery = sql

Set DTPicker1(0).DataSource = MyDDE.ActiveRecordset
Set DTPicker1(1).DataSource = MyDDE.ActiveRecordset
Set ComboStatus.DataSource = MyDDE.ActiveRecordset

DTPicker1(0).Enabled = False
DTPicker1(1).Enabled = False
End Sub

Private Function IndexAuto() As String
On Error GoTo 1
Dim Rc As New DBQuick
Dim TglSaiki As String
Dim Inom As Long
TglSaiki = Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2)
Rc.DBOpen "SELECT MAX(RIGHT([Feedback ID], 5)) AS MaxNom FROM [Customer Feedback] WHERE (GETDATE() = { fn NOW() })", CNN, lckLockReadOnly
With Rc
     If .DBRecordset.Recordcount <> 0 Then
        Inom = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + 1
     Else
        Inom = 1
     End If
     Select Case Len(Trim(Str(Inom)))
            Case 0: IndexAuto = "FE/" & TglSaiki & "-" & Trim(Str(Inom))
            Case 1: IndexAuto = "FE/" & TglSaiki & "-" & "0000" & Trim(Str(Inom))
            Case 2: IndexAuto = "FE/" & TglSaiki & "-" & "000" & Trim(Str(Inom))
            Case 3: IndexAuto = "FE/" & TglSaiki & "-" & "00" & Trim(Str(Inom))
            Case 4: IndexAuto = "FE/" & TglSaiki & "-" & "0" & Trim(Str(Inom))
     End Select
End With
Exit Function
1:
MessageBox Err.Description, "frmcustfeedback_indexauto" & Err.Number, msgOkOnly, msgExclamation
End Function

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo 1
'On Error GoTo Hell:
Set RcPartner = New DBQuick
Select Case Index
       Case 0:
            RcPartner.DBOpen " SELECT PartnerID AS [Partner ID],CompanyName as Perusahaan, Address AS Alamat, City AS Kota, PostalCode AS [Kode Pos], Country AS Negara, Phone AS Telp FROM PartnerDB WHERE (PartnerType = 'CUSTOMER') ORDER BY PartnerID", CNN, lckLockReadOnly
       Case 1:
            RcPartner.DBOpen " SELECT [PO Order].PurchaseID, [PO Order].PartnerID, [PO Order].Kurs, [PO Order].DatePurchase, [PO Order].TermPayment, [PO Order].Taxes, [PO Order].Status, [PO Order].Periode, [PO Order].TypeTrans, [PO Order].TypeLoco, PartnerDB.CompanyName, PartnerDB.Address, PartnerDB.City, [PO Order].Account, [PO Order].Discount,[PO Order].StatusSJ,[PO Order].[Require Date], [PO Order].TermMethod, [PO Order].keterangan  FROM  [PO Order] INNER JOIN   PartnerDB ON [PO Order].PartnerID = PartnerDB.PartnerID", CNN, lckLockReadOnly
       Case 2:
            RcPartner.DBOpen "SELECT TransData.TransID AS DNID, TransData.PurchaseID, Transport.ID, Transport.Expedisi, TransData.DateTrans AS [DN DATE], TransData.[No Pol],  TransData.TypeTruck, TransData.Status, PartnerDB.CompanyName, [PO Order].DatePurchase, TransData.PartnerId, TransData.RefNotes,  [Gudang Customer].[GDG ID], [Gudang Customer].[Nama Gudang], [Gudang Customer].Alamat, Regional.[RG Name] AS Kota,[PO Order].Discount, [PO Order].Kurs, [PO Order].CurrID AS [Mata Uang]" & _
                             " FROM Regional INNER JOIN [Gudang Customer] ON Regional.RG = [Gudang Customer].RG RIGHT OUTER JOIN Transport INNER JOIN TransData ON Transport.ID = TransData.ID INNER JOIN" & _
                             " PartnerDB INNER JOIN [PO Order] ON PartnerDB.PartnerID = [PO Order].PartnerID ON TransData.PurchaseID = [PO Order].PurchaseID ON  [Gudang Customer].[GDG ID] = TransData.[GDG ID]", CNN, lckLockReadOnly
       Case 3:
             RcPartner.DBOpen " SELECT  dbo.Inventory.NoItem as [No Barang], dbo.Inventory.ItemName as [Nama Barang], dbo.Inventory.UOM as Satuan, dbo.Inventory.PriceIn as Harga" & _
                              " FROM   dbo.Inventory  " & _
                              " WHERE  (dbo.Inventory.Manufacture = '1')", CNN, lckLockReadOnly

End Select

If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 0:
            mCall.FromTagActive = "MASTER CUSTOMER"
            mCall.CaptionLink = "Customer"
           
          Case 1:
             mCall.FromTagActive = "PURCHASING"
             mCall.CaptionLink = "Purchasing"
          
          Case 2:
             mCall.FromTagActive = "DO"
             mCall.CaptionLink = "DO"
             
          Case 3:
             mCall.FromTagActive = "MASTER BARANG"
             mCall.CaptionLink = "BARANG"
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgInfo
   If MyDDE.ChildRecordset.Recordcount <> 0 Then
      MyDDE.ChildRecordset.CancelBatch adAffectCurrent
      If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
   End If
End If


'Exit Sub
'Hell:
'    Err.Clear
Exit Sub
1:
MessageBox Err.Description, "frmcustfeedback_openpartner" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyMyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

End Sub

