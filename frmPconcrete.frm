VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmPconcrete 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONCETRE PRESS"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPconcrete.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   10590
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   4905
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   1005
      EditModeReplace =   -1
      BindFormTAG     =   "CP"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
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
      Height          =   4920
      Left            =   0
      ScaleHeight     =   4890
      ScaleWidth      =   10575
      TabIndex        =   11
      Top             =   0
      Width           =   10605
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "OrderID"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   1
         Tag             =   "CP"
         Top             =   315
         Width           =   2160
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   4290
         Picture         =   "frmPconcrete.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "CP"
         Top             =   323
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   4290
         Picture         =   "frmPconcrete.frx":6BDC
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "CP"
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComCtl2.DTPicker ViewTime 
         Height          =   330
         Left            =   8250
         TabIndex        =   9
         Top             =   900
         Visible         =   0   'False
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy   HH:mm"
         Format          =   59310083
         CurrentDate     =   39524
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "keterangan"
         DataSource      =   "DDE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   2
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Tag             =   "CP"
         Top             =   3990
         Width           =   7545
      End
      Begin MSDataGridLib.DataGrid DGDETAIL 
         Bindings        =   "frmPconcrete.frx":6F66
         Height          =   2355
         Left            =   180
         TabIndex        =   8
         Top             =   1305
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   4154
         _Version        =   393216
         AllowUpdate     =   -1  'True
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   19
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "No"
            Caption         =   "No"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "dd/MM/yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Tanggal_mulai"
            Caption         =   "Tanggal Mulai"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd MMM yyyy  hh:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Tanggal_selesai"
            Caption         =   "Tanggal Selesai"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd MMM yyyy  hh:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Keterangan"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Button          =   -1  'True
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "no_ekstraksi"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   2115
         TabIndex        =   3
         Tag             =   "CP"
         Top             =   690
         Width           =   2160
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "grup"
         DataSource      =   "DDE"
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
         Left            =   7125
         TabIndex        =   6
         Tag             =   "CP"
         Top             =   675
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "tanggal_ekstraksi"
         DataSource      =   "DDE"
         Height          =   315
         Left            =   7125
         TabIndex        =   5
         Tag             =   "CP"
         Top             =   315
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   59310083
         CurrentDate     =   39365
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   8850
         TabIndex        =   16
         Top             =   4470
         Width           =   1545
      End
      Begin VB.Label labe 
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By"
         Height          =   255
         Index           =   5
         Left            =   7890
         TabIndex        =   17
         Top             =   4530
         Width           =   2055
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   9915
         X2              =   7875
         Y1              =   4785
         Y2              =   4785
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Manufacture Order"
         Height          =   255
         Index           =   4
         Left            =   225
         TabIndex        =   15
         Top             =   360
         Width           =   2055
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   2250
         X2              =   210
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   3750
         Width           =   1275
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   7245
         X2              =   5205
         Y1              =   975
         Y2              =   975
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   7245
         X2              =   5205
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2250
         X2              =   210
         Y1              =   990
         Y2              =   990
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         Height          =   255
         Index           =   2
         Left            =   5220
         TabIndex        =   14
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   255
         Index           =   1
         Left            =   5205
         TabIndex        =   13
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "No Ekstraksi"
         Height          =   255
         Index           =   0
         Left            =   225
         TabIndex        =   12
         Top             =   735
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmPconcrete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Dim RsDetail As New DBQuick
Dim RsLookup As New DBQuick
Dim SelCell As Integer
Dim MEdit As Boolean
Private lMode As Byte
Private curCol As Integer

Public Property Let SetMode(ByVal Value As Byte)
   lMode = Value
End Property


Private Sub cmdLink_Click(Index As Integer)
   DoLookup Index
End Sub

Private Sub DoLookup(idx As Integer)
   Select Case idx
      Case 0
         RsLookup.DBOpen "select OrderID,OrderName,Type,RequireDate as [Tanggal Kebutuhan] from [Manufacture Order] where status='RELEASED'", CNN
         mCall.FromTagActive = "Manufacture Order"
      Case 1
         RsLookup.DBOpen "select NoEkstraksi from statusProduksi where posisi='PEMBUNGKUSAN' AND status=1", CNN
         mCall.FromTagActive = "No Ekstraksi"
   End Select
   Set mCall.FormData = RsLookup.DBRecordset
End Sub

Private Sub DDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim aNo As Integer
   cmdLink(0).Enabled = False
   cmdLink(1).Enabled = False
Select Case AdReasonActiveDb
    Case tmbAddNew:
        cmdLink(0).Enabled = True
        cmdLink(1).Enabled = True
        txt(2).Text = "-"
        txt(3).Text = frmProduksi.txtBox(5).Text
        'txt(0).Text = frmProduksi.txtBox(1).Text
        tgl.Value = Now
    Case tmbDetail:
        With DDE.ChildRecordset
            .Fields("Tanggal_Mulai") = Now
            .Fields("Tanggal_Selesai") = Now
            .Fields("keterangan") = "-"
            .MoveFirst
            aNo = 1
            While Not .EOF
                .Fields("no") = aNo
                .MoveNext
                aNo = aNo + 1
            Wend
        End With
        
    Case tmbSave:
        simpan_detail
        SaveStatusProduksi
        SaveToMO
        
    Case tmbEdit:
        cmdLink(0).Enabled = True
        cmdLink(1).Enabled = True
    
    Case tmbPrint:
        Dim lPrint As New utility
        lPrint.CallReportView "select * from concrete_press where no_ekstraksi ='" & txt(0).Text & "'", "Concrete Press.rpt", ReportPath, "Concrete Press"
        Set lPrint = Nothing
End Select
End Sub

Private Sub SaveToMO()
   Dim dStart As Date
   Dim dFinish As Date
   Dim ActualTime As Double
   Dim sWCID As String
   Dim rsCek As New DBQuick
   
   DDE.ChildRecordset.MoveFirst
   dStart = DDE.ChildRecordset.Fields("Tanggal_mulai")
   DDE.ChildRecordset.MoveLast
   dFinish = DDE.ChildRecordset.Fields("Tanggal_selesai")
   ActualTime = Val(SelisihHariJam(dStart, dFinish, 2))
   
   rsCek.DBOpen "select WCID from WCenter_Header where FormID = 45", CNN
   If rsCek.DBRecordset.Recordcount > 0 Then
      sWCID = rsCek.DBRecordset.Fields(0)
      SendDataToServer "update [order output detail] set actual_time=" & ActualTime & " where OrderID='" & DDE.GetFieldByName("OrderID") & "' and WCID='" & sWCID & "'"
   End If
   rsCek.CloseDB
End Sub

Private Sub SaveStatusProduksi()
   If Not MEdit Then
      SendDataToServer "update StatusProduksi set status=1, Posisi='CONCRETE PRESS' where noEkstraksi='" & txt(0).Text & "'"
   End If
End Sub

Private Sub DDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
   PrepareSQL
   Select Case AdReasonActiveDb
      Case tmbDelete:
         SendDataToServer "Update StatusProduksi set POsisi ='PEMBUNGKUSAN',status=1 where NoEkstraksi='" & txt(0).Text & "'"
   End Select
End Sub

Private Sub DDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    RsDetail.DBOpen "select No,Tanggal_mulai,tanggal_selesai,keterangan from CONCrete_detail where no_ekstraksi = '" & DDE.GetFieldByName("No_ekstraksi") & "'", CNN
    Set DDE.ChildRecordset = RsDetail.DBRecordset '.Clone(adLockBatchOptimistic)
    Set dgDetail.DataSource = DDE.ChildRecordset
    Label1.Caption = IIf(IsNull(DDE.GetFieldByName("approved_by")), "", DDE.GetFieldByName("approved_by"))
End Sub

Private Sub DDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim rsCek As New DBQuick

Select Case AdReasonActiveDb
   Case tmbSave:
      DDE.IsChildMemberReady = True
   Case tmbDelete:
      rsCek.DBOpen "select posisi from statusProduksi where NoEkstraksi='" & txt(0).Text & "'", CNN
      If rsCek.DBRecordset.Recordcount > 0 Then
         If rsCek.DBRecordset.Fields(0) <> "CONCRETE PRESS" Then
            MessageBox "Data Tidak Bisa dihapus Karena sedang di proses di lokasi lain", "Peringatan", msgOkOnly, msgCrtical
            DDE.CancelTrans = True
         End If
      Else
         MessageBox "Data Tidak Boleh dihapus", "Error Aplikasi", msgOkOnly, msgCrtical
         DDE.CancelTrans = True
      End If
   Case tmbAddNew: MEdit = False
   Case tmbEdit: MEdit = True
End Select
rsCek.CloseDB
End Sub


Private Sub DGDETAIL_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo Hell
   ViewTime.Visible = False
   Select Case dgDetail.col
      Case 1
         ViewTime.Value = IIf(IsNull(DDE.ChildRecordset.Fields("tanggal_Mulai")), Now, DDE.ChildRecordset.Fields("tanggal_Mulai"))
      Case 2
         ViewTime.Value = IIf(IsNull(DDE.ChildRecordset.Fields("Tanggal_Selesai")), Now, DDE.ChildRecordset.Fields("Tanggal_Selesai"))
   End Select
    
    curCol = dgDetail.col
    If (dgDetail.col = 1) Or (dgDetail.col = 2) Then
        ViewTime.Visible = True
        ViewTime.Move dgDetail.Columns(dgDetail.col).Left + dgDetail.Left, _
                      dgDetail.RowTop(dgDetail.row) + dgDetail.Top, _
                      dgDetail.Columns(dgDetail.col).width, _
                      dgDetail.RowHeight
    End If
Exit Sub
Hell:
   If Err.Number = 380 Then
      Err.Clear
      ViewTime.Value = Now
   Else
      MessageBox Err.Description, "Error", msgOkOnly, msgExclamation
   End If
End Sub

Private Sub Form_Load()

   If lMode = 0 Then
      DDE.SetReadOnlyMode = False
   Else
      DDE.SetReadOnlyMode = True
   End If

With DDE
Set .BindForm = Me
    .BindFormTAG = "CP"
Set .ActiveConnection = CNN
   If lMode = 0 Then
    .PrepareQuery = "SELECT concrete_header.no_ekstraksi, concrete_header.OrderID, concrete_header.tanggal_ekstraksi, " & _
                    "concrete_header.grup, concrete_header.Keterangan, approved_by FROM concrete_header INNER JOIN " & _
                    "[Manufacture Order] ON concrete_header.OrderID = [Manufacture Order].OrderID " & _
                    "where [manufacture order].status='RELEASED'"
   Else
    .PrepareQuery = "SELECT concrete_header.no_ekstraksi, concrete_header.OrderID, concrete_header.tanggal_ekstraksi, " & _
                    "concrete_header.grup, concrete_header.Keterangan FROM concrete_header INNER JOIN " & _
                    "[Manufacture Order] ON concrete_header.OrderID = [Manufacture Order].OrderID " & _
                    "where [manufacture order].status <> 'RELEASED'"
   End If

                    

    'If .ActiveRecordset.Recordcount > 0 Then .ActiveRecordset.MoveFirst
End With
HiasFormManTell Picture2, Me
dgDetail.RowHeight = 300
Set mCall = New frmCaller
End Sub

Function PrepareSQL()
   DDE.PrepareAppend = "insert into CONCRETE_HEADER (no_ekstraksi,OrderID, tanggal_ekstraksi, grup,keterangan,issued_by) values " & _
                       " ('" & txt(0).Text & "','" & txt(3).Text & "','" & Format(tgl.Value, "yyyy-MM-dd") & "', " & _
                       " '" & DDE.GetFieldByName("grup") & "','" & DDE.GetFieldByName("keterangan") & "','" & MainMenu.StatusBar1.Panels(1).Text & "')"
   Debug.Print DDE.PrepareAppend
   DDE.PrepareUpdate = "update CONCRETE_HEADER set tanggal_ekstraksi = '" & Format(tgl.Value, "yyyy-MM-dd") & "',OrderID='" & DDE.GetFieldByName("OrderID") & "', grup = '" & txt(1).Text & "', keterangan = '" & txt(2).Text & "' where no_ekstraksi = '" & txt(0).Text & "'"
   
   DDE.PrepareDelete = "delete from CONCRETE_HEADER where no_ekstraksi = '" & txt(0).Text & "'"
   
End Function

Function simpan_detail()
With DDE.ChildRecordset
   If .Recordcount <> 0 Then
       .MoveFirst
         If SendDataToServer(" delete from [CONCRETE_DETAIL] where (no_ekstraksi = '" & txt(0).Text & "')") = True Then
         Do
           If .EOF = True Then Exit Do
           SendDataToServer "insert into CONCRETE_DETAIL (no_ekstraksi,No,tanggal_mulai,tanggal_selesai, keterangan) " & _
           " values ('" & txt(0).Text & "', " & _
           " '" & .Fields("No") & "', " & _
           " '" & Format(.Fields("tanggal_mulai"), "yyyy-MM-dd hh:mm:ss") & "', " & _
           " '" & Format(.Fields("tanggal_selesai"), "yyyy-MM-dd hh:mm:ss") & "'," & _
           " '" & .Fields("keterangan") & "')"
          .MoveNext
        Loop
        End If
        .MoveLast
        dgDetail.Refresh
        End If
    End With
End Function



Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
   Select Case UCase(mCall.FromTagActive)
      Case "MANUFACTURE ORDER"
         txt(3).Text = mCall.GetFieldByName(0)
      Case "NO EKSTRAKSI"
         txt(0).Text = mCall.GetFieldByName(0)
   End Select
End Sub


Private Sub txt_LostFocus(Index As Integer)
   Dim rsCek As New DBQuick
   rsCek.DBOpen "select * from statusProduksi where noEkstraksi='" & txt(0).Text & "'", CNN, lckLockBatch
   If rsCek.DBRecordset.Recordcount > 0 Then
      rsCek.DBOpen "select * from concrete_header where no_Ekstraksi='" & txt(0).Text & "'", CNN, lckLockBatch
      If rsCek.DBRecordset.Recordcount > 0 Then
         MessageBox "Nomor Ekstraksi Ini Sudah Diinput...!", "Peringatan", msgOkOnly, msgCrtical
         txt(0).Text = ""
      End If
   Else
      MessageBox "Nomor Ekstraksi Ini tidak ditemukan...!", "Peringatan", msgOkOnly, msgCrtical
      txt(0).Text = ""
   End If
   rsCek.CloseDB
End Sub

Private Sub ViewTime_Change()
   Select Case curCol
      Case 1: DDE.ChildRecordset.Fields("Tanggal_Mulai") = ViewTime.Value
      Case 2: DDE.ChildRecordset.Fields("tanggal_selesai") = ViewTime.Value
   End Select
End Sub
