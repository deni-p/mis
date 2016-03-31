VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmPCrusher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crusher"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPCrusher.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   10380
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
      Height          =   4575
      Left            =   0
      ScaleHeight     =   4575
      ScaleWidth      =   10380
      TabIndex        =   5
      Top             =   0
      Width           =   10380
      Begin MSComCtl2.DTPicker viewTime 
         Height          =   300
         Left            =   2910
         TabIndex        =   9
         Top             =   1770
         Visible         =   0   'False
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy  HH:mm "
         Format          =   63111171
         CurrentDate     =   39530
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
         Height          =   330
         Index           =   1
         Left            =   8010
         TabIndex        =   3
         Tag             =   "cruz"
         Top             =   540
         Width           =   2175
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "no_crusher"
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
         Height          =   330
         Index           =   0
         Left            =   825
         TabIndex        =   1
         Tag             =   "cruz"
         Top             =   210
         Width           =   2190
      End
      Begin MSDataGridLib.DataGrid DGDETAIL 
         Bindings        =   "frmPCrusher.frx":6852
         Height          =   3255
         Left            =   90
         TabIndex        =   4
         Top             =   930
         Width           =   10170
         _ExtentX        =   17939
         _ExtentY        =   5741
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "no_ekstraksi"
            Caption         =   "No Ekstraksi"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "kondisi"
            Caption         =   "Kondisi Mesin"
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
            DataField       =   "Tanggal_mulai"
            Caption         =   "Mulai Tanggal"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd MMM yyyy hh:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "tanggal_selesai"
            Caption         =   "Tanggal Selesai"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd MMM yyyy hh:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "kuantitas"
            Caption         =   "Kuantitas (Kg)"
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
            DataField       =   "keterangan"
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
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Button          =   -1  'True
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Button          =   -1  'True
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
            EndProperty
            BeginProperty Column05 
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "tanggal_crusher"
         DataSource      =   "DDE"
         Height          =   315
         Left            =   8010
         TabIndex        =   2
         Tag             =   "cruz"
         Top             =   180
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   63111171
         CurrentDate     =   39365
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7905
         TabIndex        =   10
         Top             =   4230
         Width           =   2340
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By"
         Height          =   255
         Index           =   4
         Left            =   6345
         TabIndex        =   12
         Top             =   4305
         Width           =   2055
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   8385
         X2              =   6345
         Y1              =   4530
         Y2              =   4530
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   8700
         X2              =   6660
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   8685
         X2              =   6645
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         Height          =   255
         Index           =   2
         Left            =   6660
         TabIndex        =   8
         Top             =   615
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   255
         Index           =   1
         Left            =   6645
         TabIndex        =   7
         Top             =   255
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   255
         Width           =   2055
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2175
         X2              =   135
         Y1              =   510
         Y2              =   510
      End
   End
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   4575
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   1005
      BindFormTAG     =   "cruz"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   2055
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2040
      X2              =   0
      Y1              =   225
      Y2              =   225
   End
End
Attribute VB_Name = "frmPCrusher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Dim RsLookup As New DBQuick
Private RsDetail As New DBQuick
Private lMode As Byte

Public Property Let SetMode(ByVal Value As Byte)
   lMode = Value
End Property


Private Sub loadDetail()
    RsDetail.DBOpen "select * from crusher_detail where no_crusher = '" & DDE.GetFieldByName("no_crusher") & "'", CNN
    Set DDE.ChildRecordset = RsDetail.DBRecordset.Clone(adLockBatchOptimistic)
    Set DGDETAIL.DataSource = DDE.ChildRecordset
End Sub

Private Sub cmdLink_Click()
   RsLookup.DBOpen "select OrderID,OrderName,Type,RequireDate from [Manufacture Order] where status='RELEASED'", CNN
   If RsLookup.DBRecordset.Recordcount > 0 Then
      Set mCall.FormData = RsLookup.DBRecordset
      mCall.FromTagActive = "Manufacture Order"
   Else
      MessageBox "Data Tidak Tersedia", "Peringatan", msgOkOnly, msgCrtical
   End If
End Sub

Private Sub DDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
   Select Case AdReasonActiveDb
      Case tmbAddNew: txt(0).Text = IndexAuto
         tgl.Value = Now
         DDE.GetFieldByName("tanggal_crusher") = Now
      Case tmbDetail:
         LoadEkstraksi
      Case tmbSave:
         simpan_detail
      Case tmbPrint:
         Dim lPrint As New utility
         lPrint.CallReportView "select * from crusher_report where no_crusher='" & txt(0).Text & "'", "crusher.rpt", ReportPath, "Crusher"
         Set lPrint = Nothing

   End Select
End Sub

Private Sub LoadEkstraksi()
     RsLookup.DBOpen "select * from statusProduksi " & SQLLookupParameter(DDE.ChildRecordset, "NoEkstraksi", "No_ekstraksi", " posisi='JEMUR' or posisi='DRYER' "), CNN
     If RsLookup.DBRecordset.Recordcount > 0 Then
         Set mCall.FormData = RsLookup.DBRecordset
         mCall.FromTagActive = "No Ekstraksi"
     Else
         DDE.ChildRecordset.MoveLast
         DDE.ChildRecordset.Delete
     End If
End Sub


Private Sub DDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
   Select Case AdReasonActiveDb
      Case tmbDelete:
         If DDE.ChildRecordset.Recordcount > 0 Then
            SendDataToServer "update statusProduksi set posisi='JEMUR',status=1 where NoEkstraksi='" & DDE.ChildRecordset("no_ekstraksi") & "'"
         End If
   End Select
End Sub

Private Sub DDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   If IsEmpty(DDE.GetFieldByName("tanggal_crusher")) Or IsNull(DDE.GetFieldByName("tanggal_crusher")) Then
      tgl.Value = Now
   Else
      tgl.Value = DDE.GetFieldByName("tanggal_crusher")
   End If
    loadDetail
    Label1.Caption = IIf(IsNull(DDE.GetFieldByName("Approved_by")), "", DDE.GetFieldByName("Approved_by"))
End Sub

Function PrepareSQL()
   DDE.PrepareAppend = "insert into CRUSHER_HEADER (no_crusher,tanggal_crusher,grup,issued_by) values " & _
                        " ('" & txt(0).Text & "','" & Format(tgl.Value, "yyyy-MM-dd") & "','" & DDE.GetFieldByName("grup") & "','" & MainMenu.StatusBar1.Panels(1).Text & "')"
    
   DDE.PrepareUpdate = "update CRUSHER_HEADER set tanggal_crusher = '" & Format(tgl.Value, "yyyy-MM-dd") & "', grup = '" & DDE.GetFieldByName("grup") & "' where no_crusher = '" & txt(0).Text & "'"

   DDE.PrepareDelete = "delete from CRUSHER_HEADER where tanggal_crusher = '" & Format(tgl.Value, "yyyy-MM-dd") & "'"

End Function

Function simpan_detail()
With DDE.ChildRecordset
   If .Recordcount <> 0 Then
       .MoveFirst
         If SendDataToServer(" delete from [CRUSHER_DETAIL] where (no_crusher = '" & DDE.GetFieldByName("no_crusher") & "')") = True Then
         Do
           If .EOF = True Then Exit Do
           SendDataToServer "insert into CRUSHER_DETAIL (no_crusher,no_ekstraksi,kondisi,tanggal_mulai, tanggal_selesai, kuantitas,keterangan)  " & _
           " values ('" & txt(0).Text & "','" & DGDETAIL.Columns(0) & "', '" & .Fields("kondisi") & "'," & _
           " '" & Format(.Fields("tanggal_mulai"), "yyyy-MM-dd hh:mm:ss") & "', " & _
           " '" & Format(.Fields("tanggal_selesai"), "yyyy-MM-dd hh:mm:ss") & "'," & _
           " '" & .Fields("kuantitas") & _
           "','" & .Fields("keterangan") & "')"
            
           SendDataToServer "update StatusProduksi set status=1, posisi='CRUSHER',tanggal='" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "' where noEkstraksi='" & DGDETAIL.Columns(0) & "' "
           
           SaveToMO
          .MoveNext
        Loop
        End If
        .MoveLast
        DGDETAIL.Refresh
        End If
    End With
End Function
Private Sub DDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
   PrepareSQL
   Select Case AdReasonActiveDb
      Case tmbAddNew:
      Case tmbEdit:
      Case tmbSave:
           DDE.IsChildMemberReady = True
   End Select
End Sub


Private Sub DGDETAIL_KeyPress(KeyAscii As Integer)
Dim No As Integer
If KeyAscii = 13 Then
   DDE.ChildRecordset.AddNew
End If
End Sub

Private Sub DGDETAIL_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo Hell
   viewTime.Visible = False
   Select Case DGDETAIL.col
      Case 2: viewTime.Visible = True
              viewTime.Value = IIf(IsNull(DDE.ChildRecordset.Fields("tanggal_mulai")), Now, DDE.ChildRecordset.Fields("tanggal_mulai"))
              viewTime.Move DGDETAIL.Left + DGDETAIL.Columns(2).Left, _
                            DGDETAIL.Top + DGDETAIL.RowTop(DGDETAIL.row), _
                            DGDETAIL.Columns(2).width, _
                            DGDETAIL.RowHeight
      Case 3: viewTime.Visible = True
              viewTime.Value = IIf(IsNull(DDE.ChildRecordset.Fields("tanggal_selesai")), Now, DDE.ChildRecordset.Fields("tanggal_selesai"))
              viewTime.Move DGDETAIL.Left + DGDETAIL.Columns(3).Left, _
                            DGDETAIL.Top + DGDETAIL.RowTop(DGDETAIL.row), _
                            DGDETAIL.Columns(3).width, _
                            DGDETAIL.RowHeight
   End Select
Exit Sub
Hell:
   If Err.Number = 380 Then
      Err.Clear
   Else
      MessageBox Err.Description, "Error", msgOkOnly, msgExclamation
   End If
End Sub


Private Function IndexAuto() As String
Dim Rc As New DBQuick
Dim TglSaiki As String
Dim Inom As Long
TglSaiki = Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2)
Rc.DBOpen "SELECT MAX(RIGHT(No_crusher, 5)) AS MaxNom FROM [crusher_header] WHERE (GETDATE() = { fn NOW() })", CNN, lckLockReadOnly
With Rc
     If .DBRecordset.Recordcount <> 0 Then
        Inom = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + 1
     Else
        Inom = 1
     End If
     Select Case Len(Trim(Str(Inom)))
            Case 0: IndexAuto = "CR/" & TglSaiki & "-" & Trim(Str(Inom))
            Case 1: IndexAuto = "CR/" & TglSaiki & "-" & "0000" & Trim(Str(Inom))
            Case 2: IndexAuto = "CR/" & TglSaiki & "-" & "000" & Trim(Str(Inom))
            Case 3: IndexAuto = "CR/" & TglSaiki & "-" & "00" & Trim(Str(Inom))
            Case 4: IndexAuto = "CR/" & TglSaiki & "-" & "0" & Trim(Str(Inom))
     End Select
End With
End Function
Private Sub Form_Load()

   If lMode = 0 Then
      DDE.SetReadOnlyMode = False
   Else
      DDE.SetReadOnlyMode = True
   End If

   With DDE
   Set .BindForm = Me
       .BindFormTAG = "cruz"
   Set .ActiveConnection = CNN
       .PrepareQuery = "select * from CRUSHER_HEADER order by tanggal_crusher desc"
   End With
   HiasFormManTell Picture2, Me
   Set mCall = New frmCaller
   DGDETAIL.RowHeight = 300
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
   If mCall.FromTagActive = "Manufacture Order" Then
      DDE.GetFieldByName("OrderId") = mCall.GetFieldByName(0)
   Else
      If RsLookup.DBRecordset.Recordcount > 0 Then
         DDE.ChildRecordset.MoveLast
         DDE.ChildRecordset.Fields("no_ekstraksi") = mCall.GetFieldByName(0)
         DDE.ChildRecordset.Fields("tanggal_mulai") = Now
         DDE.ChildRecordset.Fields("tanggal_selesai") = Now
      End If
   End If
End Sub


Private Sub ViewTime_Change()
   Select Case DGDETAIL.col
      Case 2: DDE.ChildRecordset.Fields("tanggal_mulai") = viewTime.Value
      Case 3: DDE.ChildRecordset.Fields("tanggal_selesai") = viewTime.Value
   End Select
End Sub


Private Sub SaveToMO()
   Dim dStart As Date
   Dim dFinish As Date
   Dim ActualTime As Double
   Dim sWCID As String
   Dim rsCek As New DBQuick
   
   dStart = DDE.ChildRecordset.Fields("Tanggal_mulai")
   dFinish = DDE.ChildRecordset.Fields("Tanggal_selesai")
   ActualTime = Val(SelisihHariJam(dStart, dFinish, 2))
   
   rsCek.DBOpen "select WCID from WCenter_Header where FormID = 48", CNN
   If rsCek.DBRecordset.Recordcount > 0 Then
      sWCID = rsCek.DBRecordset.Fields(0)
      SendDataToServer "update [order output detail] set actual_time=" & ActualTime & " where OrderID='" & DDE.GetFieldByName("OrderID") & "' and WCID='" & sWCID & "'"
   End If
   rsCek.CloseDB
End Sub

