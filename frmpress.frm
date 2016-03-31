VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmpress 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HYDRAULIC PRESS"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12165
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmpress.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   12165
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   4770
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   1005
      BindFormTAG     =   "AA"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   0
      ScaleHeight     =   4770
      ScaleWidth      =   12165
      TabIndex        =   7
      Top             =   0
      Width           =   12165
      Begin MSComCtl2.DTPicker ViewDate 
         Height          =   315
         Left            =   9660
         TabIndex        =   12
         Top             =   3660
         Visible         =   0   'False
         Width           =   1890
         _ExtentX        =   3334
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
         CustomFormat    =   "dd MMM yyyy   HH:mm"
         Format          =   63242243
         CurrentDate     =   39525
      End
      Begin MSComCtl2.DTPicker ViewTime 
         Height          =   315
         Left            =   8460
         TabIndex        =   11
         Top             =   3675
         Visible         =   0   'False
         Width           =   1110
         _ExtentX        =   1958
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
         Format          =   63242242
         CurrentDate     =   39525
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "no_press"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   2
         Left            =   2295
         TabIndex        =   1
         Tag             =   "AA"
         Top             =   195
         Width           =   1695
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "keterangan"
         DataSource      =   "DDE"
         Height          =   780
         Index           =   0
         Left            =   105
         MultiLine       =   -1  'True
         TabIndex        =   4
         Tag             =   "AA"
         Text            =   "frmpress.frx":6852
         Top             =   3900
         Width           =   5985
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "grup"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   1
         Left            =   7335
         TabIndex        =   3
         Tag             =   "AA"
         Top             =   555
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "tanggal_press"
         DataSource      =   "DDE"
         Height          =   315
         Left            =   7335
         TabIndex        =   2
         Tag             =   "AA"
         Top             =   195
         Width           =   2175
         _ExtentX        =   3836
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
         CustomFormat    =   "dd MMM yyyy"
         Format          =   63242243
         CurrentDate     =   39365
      End
      Begin MSDataGridLib.DataGrid DGDETAIL 
         Height          =   2625
         Left            =   120
         TabIndex        =   5
         Top             =   990
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   4630
         _Version        =   393216
         AllowUpdate     =   -1  'True
         DefColWidth     =   1
         HeadLines       =   3
         RowHeight       =   21
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "no_ekstraksi"
            Caption         =   "No Ekstraksi"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "no_kereta"
            Caption         =   "No Kereta"
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
            DataField       =   "no_hyd_press"
            Caption         =   "No Hyd. Press"
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
            DataField       =   "jml_stack"
            Caption         =   "Jml Stack"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "no_stack"
            Caption         =   "No Stack"
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
            DataField       =   "ready"
            Caption         =   "Waktu Siap Dikereta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "HH:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "tanggal_mulai"
            Caption         =   "Tgl Mulai"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd MMM yyyy HH:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "tanggal_selesai"
            Caption         =   "Tgl Selesai"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd MMM yyyy HH:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "tekanan_max"
            Caption         =   "Tekanan (Bar)"
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
               Alignment       =   3
               WrapText        =   -1  'True
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               WrapText        =   -1  'True
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column08 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   9450
         TabIndex        =   13
         Top             =   4320
         Width           =   2580
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By"
         Height          =   255
         Index           =   4
         Left            =   7530
         TabIndex        =   14
         Top             =   4395
         Width           =   2055
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   9555
         X2              =   7515
         Y1              =   4650
         Y2              =   4650
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2400
         X2              =   360
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Top             =   3660
         Width           =   855
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   255
         Index           =   1
         Left            =   5415
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         Height          =   255
         Index           =   2
         Left            =   5430
         TabIndex        =   8
         Top             =   630
         Width           =   2055
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   7455
         X2              =   5415
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   7470
         X2              =   5430
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "No Hydraulic Press"
         Height          =   255
         Index           =   3
         Left            =   375
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmpress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IDGen As New IDGenerator
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Dim RsDetail As New DBQuick
Private RsLookup As New DBQuick
Private MEdit As Boolean
Private lMode As Byte

Public Property Let SetMode(ByVal Value As Byte)
   lMode = Value
End Property


Private Sub ReverseStatusOrder(noEkstraksi As String)
   Dim rsCek As New DBQuick
   rsCek.DBOpen "select posisi from statusProduksi where noEkstraksi ='" & noEkstraksi & "'", CNN
   If rsCek.DBRecordset.Recordcount > 0 Then
      If rsCek.DBRecordset.Fields(0) = "HYDRAULIC PRESS" Then UpdateStatusProduksi "CONCRETE PRESS", noEkstraksi
   End If
   rsCek.CloseDB
End Sub

Private Sub UpdateStatusProduksi(strStatus As String, noEkstraksi As String)
   SendDataToServer "Update statusProduksi set posisi='" & strStatus & "',status=1 where NoEkstraksi = '" & noEkstraksi & "'"
End Sub

Private Sub cmd_Click()
   DoLookup 0
End Sub


Private Sub DoLookup(idx As Integer)
   Select Case idx
      Case 0: RsLookup.DBOpen "select OrderID,OrderName,Type,RequireDate from [Manufacture Order] where status='RELEASED'", CNN
      Case 1: RsLookup.DBOpen "select noEkstraksi  from statusProduksi where posisi='CONCRETE PRESS' and status=1 ", CNN
   End Select
   If RsLookup.Recordcount > 0 Then
      Set mCall.FormData = RsLookup.DBRecordset
      Select Case idx
         Case 0: mCall.FromTagActive = "Manufacture Order"
         Case 1: mCall.FromTagActive = "No Ekstraksi"
      End Select
   Else
      MessageBox "Data Tidak Tersedia", "Peringatan", msgOkOnly, msgCrtical
      DDE.ChildRecordset.MoveLast
      DDE.ChildRecordset.Delete
   End If
End Sub


Private Sub DDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
    Case tmbAddNew:
        DDE.GetFieldByName("No_press") = IDGen.GetID("HP")
        DDE.GetFieldByName("grup") = "-"
        DDE.GetFieldByName("keterangan") = "-"
        DDE.GetFieldByName("tanggal_press") = Now
        DDE.GetFieldByName("issued_by") = MainMenu.StatusBar1.Panels(1).Text
        tgl.Value = Now
    Case tmbSave:
      If DDE.IsChildMemberReady Then simpan_detail
        
    Case tmbDetail:
        DoLookup 1
    Case tmbPrint:
      Dim lPrint As New utility
     ' Debug.Print "select * from hydraulic_press where no_hyd_press ='" & txt(2).Text & "'"
      
      lPrint.CallReportView "select * from hydraulic_press where no_press ='" & txt(2).Text & "'", "Hydraulic Press.rpt", ReportPath, "Hydraulic Press"
      Set lPrint = Nothing
End Select

End Sub

Private Sub DDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    loadDetail
    Label1.Caption = IIf(IsNull(DDE.GetFieldByName("Approved_by")), "", DDE.GetFieldByName("Approved_by"))
End Sub

Private Sub loadDetail()
    RsDetail.DBOpen "select no_ekstraksi,no_kereta,no_hyd_press,jml_stack,no_stack,ready,tanggal_mulai,tanggal_selesai,tekanan_max from press_detail where NO_PRESS = '" & DDE.GetFieldByName("NO_PRESS") & "'", CNN
    Set DDE.ChildRecordset = RsDetail.DBRecordset '.Clone(adLockBatchOptimistic)
    Set dgDetail.DataSource = DDE.ChildRecordset
End Sub

Private Sub DDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareSQL
Select Case AdReasonActiveDb
    Case tmbAddNew:
         MEdit = False
    Case tmbEdit:
         MEdit = True
    Case tmbSave: DDE.IsChildMemberReady = True
    Case tmbDelete:
      If DDE.ChildRecordset.Recordcount > 0 Then
         ReverseStatusOrder DDE.ChildRecordset.Fields("No_ekstraksi")
      End If
End Select
End Sub


Private Sub MoveObj(oObj As Object)
    oObj.Visible = True
    oObj.Move dgDetail.Columns(dgDetail.col).Left + dgDetail.Left, _
                          dgDetail.RowTop(dgDetail.row) + dgDetail.Top, _
                          dgDetail.Columns(dgDetail.col).width, _
                          dgDetail.RowHeight
    oObj.SetFocus
End Sub

Private Sub DGDETAIL_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo Hell
    viewTime.Visible = False
    viewDate.Visible = False
    Select Case dgDetail.col
        Case 5:
            viewTime.Value = IIf(IsNull(DDE.ChildRecordset.Fields("ready")), Now, DDE.ChildRecordset.Fields("ready"))
            MoveObj viewTime
            viewTime.Visible = True
        Case 6:
            viewDate.Value = DDE.ChildRecordset.Fields("tanggal_mulai")
            MoveObj viewDate
            viewDate.Visible = True
        Case 7:
            viewDate.Value = DDE.ChildRecordset.Fields("tanggal_selesai")
            MoveObj viewDate
            viewDate.Visible = True
    End Select
Exit Sub
Hell:
   If Err.Number = 380 Then
      Err.Clear
      viewDate.Value = Now
      viewTime.Value = Now
   Else
      MessageBox Err.Description, "Peringatan", msgOkOnly, msgExclamation
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
      .BindFormTAG = "AA"
      Set .ActiveConnection = CNN
          .PrepareQuery = "select * from PRESS_HEADER"
          If .ActiveRecordset.Recordcount > 0 Then .ActiveRecordset.MoveFirst
          loadDetail
   End With
   Set mCall = New frmCaller
   HiasFormManTell Picture2, Me
   dgDetail.RowHeight = 300
   dgDetail.HeadLines = 3
End Sub

Function PrepareSQL()
   With DDE
      .PrepareAppend = "insert into PRESS_HEADER(no_press,tanggal_press,grup,keterangan,issued_by) values " & _
                          "('" & .GetFieldByName("no_press") & "','" & _
                          Format(tgl.Value, "yyyy-MM-dd") & "', '" & _
                          .GetFieldByName("grup") & "', '" & _
                          DDE.GetFieldByName("keterangan") & "','" & _
                          MainMenu.StatusBar1.Panels(1).Text & "')"
      
      .PrepareUpdate = "update PRESS_HEADER set tanggal_press = '" & Format(tgl.Value, "yyyy-MM-dd") & "'," & _
                                                "grup = '" & DDE.GetFieldByName("grup") & "'," & _
                                                "keterangan = '" & DDE.GetFieldByName("keterangan") & "'," & _
                                                "issued_by = '" & MainMenu.StatusBar1.Panels(1).Text & "' " & _
                        " where no_press = '" & txt(2).Text & "'"
      
      .PrepareDelete = "Delete from Press_header where no_press ='" & txt(2).Text & "'"
   End With
End Function

Function simpan_detail()
With DDE.ChildRecordset
   If .Recordcount <> 0 Then
       .MoveFirst
         If SendDataToServer(" delete from [PRESS_DETAIL] where (NO_PRESS= '" & DDE.GetFieldByName("NO_PRESS") & "')") = True Then
         Do
           If .EOF = True Then Exit Do
           ReverseStatusOrder .Fields("no_ekstraksi")
           
           SendDataToServer "insert into PRESS_DETAIL (NO_PRESS,no_ekstraksi,no_kereta,no_hyd_press,jml_stack, no_stack,ready,tanggal_mulai, tanggal_selesai, tekanan_max)  " & _
           " values ('" & DDE.GetFieldByName("NO_PRESS") & "','" & .Fields("no_ekstraksi") & "', " & _
           " '" & .Fields("no_kereta") & "', " & _
           " '" & .Fields("no_hyd_press") & "', " & _
           " '" & .Fields("jml_stack") & "', " & _
           " '" & .Fields("no_stack") & "'," & _
           " '" & Format(.Fields("ready"), "yyyy-MM-dd hh:mm:ss") & "', " & _
           " '" & Format(.Fields("tanggal_mulai"), "yyyy-MM-dd hh:mm:ss") & "'," & _
           " '" & Format(.Fields("tanggal_selesai"), "yyyy-MM-dd hh:mm:ss") & "'," & _
           " '" & .Fields("tekanan_max") & "')"
           
           UpdateStatusProduksi "HYDRAULIC PRESS", .Fields("No_ekstraksi")
          .MoveNext
        Loop
        End If
        .MoveLast
        dgDetail.Refresh
        End If
    End With
End Function


Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
   Select Case UCase(mCall.FromTagActive)
      Case "MANUFACTURE ORDER"
         DDE.GetFieldByName("OrderID") = mCall.GetFieldByName(0)
      Case "NO EKSTRAKSI"
         If pRecordset.Recordcount > 0 Then
            DDE.ChildRecordset.MoveLast
            DDE.ChildRecordset.Fields("no_ekstraksi") = mCall.GetFieldByName(0)
            DDE.ChildRecordset.Fields("Ready") = Now
            DDE.ChildRecordset.Fields("tanggal_mulai") = Now
            DDE.ChildRecordset.Fields("tanggal_selesai") = Now
         Else
            DDE.ChildRecordset.MoveLast
            DDE.ChildRecordset.Delete
         End If

   End Select
End Sub


Private Sub viewDate_Change()
   Select Case dgDetail.col
      Case 6: DDE.ChildRecordset.Fields("tanggal_mulai") = viewDate.Value
      Case 7: DDE.ChildRecordset.Fields("tanggal_selesai") = viewDate.Value
   End Select
End Sub

Private Sub ViewTime_Change()
On Error Resume Next
   DDE.ChildRecordset.Fields("ready") = viewTime.Value
End Sub

Private Sub SaveToMO()
   Dim dStart As Date
   Dim dFinish As Date
   Dim ActualTime As Double
   Dim sWCID As String
   Dim sMO As String
   Dim rsCek As New DBQuick
   
   DDE.ChildRecordset.MoveFirst
   dStart = DDE.ChildRecordset.Fields("Tanggal_mulai")
   DDE.ChildRecordset.MoveLast
   dFinish = DDE.ChildRecordset.Fields("Tanggal_selesai")
   ActualTime = Val(SelisihHariJam(dStart, dFinish, 2))
   
   rsCek.DBOpen "select WCID from WCenter_Header where FormID = 45", CNN
   If rsCek.DBRecordset.Recordcount > 0 Then
      sWCID = rsCek.DBRecordset.Fields(0)
      rsCek.DBOpen "select OrderID from [manufacture Order] where ekstraksi_no='" & DDE.ChildRecordset("no_ekstraksi") & "'", CNN
      If rsCek.DBRecordset.Recordcount > 0 Then sMO = rsCek.DBRecordset.Fields(0) Else sMO = ""
      SendDataToServer "update [order output detail] set actual_time=" & ActualTime & " where OrderID='" & sMO & "' and WCID='" & sWCID & "'"
   End If
   rsCek.CloseDB
End Sub

