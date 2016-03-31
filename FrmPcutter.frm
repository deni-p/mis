VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmPcutter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cutter"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPcutter.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   8955
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
      Height          =   3915
      Left            =   0
      ScaleHeight     =   3915
      ScaleWidth      =   8955
      TabIndex        =   18
      Top             =   0
      Width           =   8955
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "OrderID"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   1
         Tag             =   "cutter"
         Top             =   360
         Width           =   1710
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   4290
         Picture         =   "FrmPcutter.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "cutter"
         Top             =   368
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   4290
         Picture         =   "FrmPcutter.frx":6BDC
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "cutter"
         Top             =   728
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Kondisi Lembaran Agar"
         Height          =   645
         Left            =   5415
         TabIndex        =   17
         Top             =   1110
         Width           =   3075
         Begin VB.OptionButton OptAgarA 
            BackColor       =   &H00EAAF6F&
            Caption         =   "Bersih"
            Height          =   285
            Left            =   165
            TabIndex        =   13
            Tag             =   "cutter"
            Top             =   270
            Value           =   -1  'True
            Width           =   1200
         End
         Begin VB.OptionButton OptAgarB 
            BackColor       =   &H00EAAF6F&
            Caption         =   "Kotor"
            Height          =   225
            Left            =   1530
            TabIndex        =   14
            Tag             =   "cutter"
            Top             =   270
            Width           =   1440
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Kondisi Mesin Cutter"
         Height          =   645
         Left            =   5415
         TabIndex        =   16
         Top             =   315
         Width           =   3075
         Begin VB.OptionButton OptMesinB 
            BackColor       =   &H00EAAF6F&
            Caption         =   "Kotor"
            Height          =   225
            Left            =   1530
            TabIndex        =   12
            Tag             =   "cutter"
            Top             =   270
            Width           =   1485
         End
         Begin VB.OptionButton OptMesinA 
            BackColor       =   &H00EAAF6F&
            Caption         =   "Bersih"
            Height          =   285
            Left            =   165
            TabIndex        =   11
            Tag             =   "cutter"
            Top             =   270
            Value           =   -1  'True
            Width           =   1200
         End
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "k_awal"
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
         Index           =   3
         Left            =   2565
         TabIndex        =   9
         Tag             =   "cutter"
         Top             =   2535
         Width           =   1695
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
         Height          =   1575
         Index           =   5
         Left            =   5445
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Tag             =   "cutter"
         Top             =   2190
         Width           =   3075
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "no_ekstraksi"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   2580
         TabIndex        =   3
         Tag             =   "cutter"
         Top             =   720
         Width           =   1710
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "k_akhir"
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
         Index           =   4
         Left            =   2565
         TabIndex        =   10
         Tag             =   "cutter"
         Top             =   2895
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker tgl_awal 
         DataField       =   "tgl_mulai"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "m/d/yy h:nn AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         DataSource      =   "DDE"
         Height          =   315
         Left            =   2580
         TabIndex        =   7
         Tag             =   "cutter"
         Top             =   1800
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy   HH:mm"
         Format          =   62783491
         CurrentDate     =   39401
      End
      Begin MSComCtl2.DTPicker tgl_cutter 
         DataField       =   "tanggal_cutter"
         DataSource      =   "DDE"
         Height          =   315
         Left            =   2580
         TabIndex        =   5
         Tag             =   "cutter"
         Top             =   1080
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   62783491
         CurrentDate     =   39365
      End
      Begin MSComCtl2.DTPicker tgl_akhir 
         DataField       =   "tgl_selesai"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "m/d/yy h:nn AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         DataSource      =   "DDE"
         Height          =   315
         Left            =   2565
         TabIndex        =   8
         Tag             =   "cutter"
         Top             =   2160
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy   HH:mm"
         Format          =   62783491
         CurrentDate     =   39401
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
         Left            =   2580
         TabIndex        =   6
         Tag             =   "cutter"
         Top             =   1455
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2565
         TabIndex        =   28
         Top             =   3420
         Width           =   1695
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   3000
         X2              =   420
         Y1              =   3735
         Y2              =   3735
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   2940
         X2              =   360
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Manufacture Order"
         Height          =   255
         Index           =   7
         Left            =   375
         TabIndex        =   27
         Top             =   405
         Width           =   2055
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   2940
         X2              =   360
         Y1              =   3195
         Y2              =   3195
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   2925
         X2              =   345
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   2925
         X2              =   345
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Kuantitas Awal                                                                Lembar"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   26
         Top             =   2580
         Width           =   5190
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Dan Waktu Selesai"
         Height          =   255
         Index           =   4
         Left            =   345
         TabIndex        =   25
         Top             =   2220
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Dan Waktu Mulai"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   24
         Top             =   1890
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   50
         Left            =   5445
         TabIndex        =   23
         Top             =   1890
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "No Ekstraksi"
         Height          =   255
         Index           =   0
         Left            =   375
         TabIndex        =   22
         Top             =   765
         Width           =   2055
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2940
         X2              =   360
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2985
         X2              =   360
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   2910
         X2              =   360
         Y1              =   1740
         Y2              =   1740
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   2925
         X2              =   345
         Y1              =   2835
         Y2              =   2835
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   20
         Top             =   1530
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   21
         Top             =   1155
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Kuantitas Akhir                                                                Lembar"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   19
         Top             =   2970
         Width           =   5190
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By"
         Height          =   255
         Index           =   8
         Left            =   435
         TabIndex        =   29
         Top             =   3495
         Width           =   3960
      End
   End
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   3930
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   1005
      BindFormTAG     =   "cutter"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmPcutter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RsLookup As New DBQuick
Private MEdit As Boolean
Private lMode As Byte

Public Property Let SetMode(ByVal Value As Byte)
   lMode = Value
End Property




Private Sub cmdLink_Click(Index As Integer)
   Select Case Index
      Case 0: RsLookup.DBOpen "select NoEkstraksi from statusProduksi where posisi='HYDRAULIC PRESS' and status=1", CNN
      Case 1: RsLookup.DBOpen "Select OrderID,OrderName,Type,RequireDate from [Manufacture Order] where status ='RELEASED'", CNN
   End Select
         
   If RsLookup.DBRecordset.Recordcount > 0 Then
      Set mCall.FormData = RsLookup.DBRecordset
      If Index = 0 Then
         mCall.FromTagActive = "No Ekstraksi"
      Else
         mCall.FromTagActive = "Manufacture Order"
      End If
   Else
      MessageBox "Data Tidak Tersedia", "Peringatan", msgOkOnly, msgCrtical
   End If
End Sub

Private Sub DDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim rsX As New DBQuick
   Select Case AdReasonActiveDb
      Case tmbAddNew:
          DDE.GetFieldByName("tanggal_cutter") = Now
          DDE.GetFieldByName("tgl_mulai") = Now
          DDE.GetFieldByName("tgl_selesai") = Now
          DDE.GetFieldByName("keterangan") = "-"
          txt(2).Text = frmProduksi.txtBox(5)
          'txt(0).Text = frmProduksi.txtBox(1)
          tgl_cutter.Value = Now
          tgl_awal.Value = Now
          tgl_akhir.Value = Now
          rsX.DBOpen "select hasil from produksi_pembungkusan where NoEkstraksi='" & txt(0).Text & "'", CNN
          If rsX.DBRecordset.Recordcount > 0 Then
            txt(3).Text = rsX.DBRecordset.Fields(0)
          Else
            txt(3).Text = 0
          End If
      Case tmbSave:
         If DDE.IsChildMemberReady = True Then
            SaveToMO
            If Not MEdit Then UpdateStatusProduksi "CUTTER", txt(0).Text
         End If
      Case tmbPrint:
         Dim lPrint As New utility
         lPrint.CallReportView "select * from cutter where no_ekstraksi='" & txt(0).Text & "'", "cutter.rpt", ReportPath, "Cutter"
         Set lPrint = Nothing
         
   End Select
End Sub

Private Sub UpdateStatusProduksi(strPosisi As String, noEkstraksi As String)
   Dim strSQL As String
   strSQL = "Update StatusProduksi set Posisi ='" & strPosisi & "',status=1 where NoEkstraksi ='" & noEkstraksi & "'"
  ' Debug.Print strSQL
   SendDataToServer strSQL
End Sub

Private Sub DDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
   Select Case AdReasonActiveDb
      Case tmbDelete:
         UpdateStatusProduksi "HYDRAULIC PRESS", txt(0).Text
   End Select
End Sub

Private Sub DDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   If DDE.ActiveRecordset.Recordcount > 0 Then
      OptMesinA.Value = IIf(IsNull(DDE.GetFieldByName("KondisiMesin")), True, DDE.GetFieldByName("KondisiMesin"))
      OptMesinB.Value = Not OptMesinA.Value
      OptAgarA.Value = IIf(IsNull(DDE.GetFieldByName("KondisiAgar")), True, DDE.GetFieldByName("KondisiAgar"))
      OptAgarB.Value = Not OptAgarA.Value
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
       .BindFormTAG = "cutter"
   Set .ActiveConnection = CNN
   
   If lMode = 0 Then
       .PrepareQuery = "SELECT CUTTER.No_ekstraksi, CUTTER.OrderID, CUTTER.Tanggal_cutter, CUTTER.Grup," & _
                        "CUTTER.Tgl_mulai, CUTTER.tgl_selesai, CUTTER.k_awal , CUTTER.k_akhir, " & _
                        "CUTTER.Keterangan, CUTTER.kondisiMesin, CUTTER.KondisiAgar, Cutter.issued_by " & _
                        " FROM CUTTER INNER JOIN " & _
                        " [Manufacture Order] ON CUTTER.OrderID = [Manufacture Order].OrderID " & _
                        " WHERE [Manufacture Order].Status = 'RELEASED'"
   Else
       .PrepareQuery = "SELECT CUTTER.No_ekstraksi, CUTTER.OrderID, CUTTER.Tanggal_cutter, CUTTER.Grup," & _
                        "CUTTER.Tgl_mulai, CUTTER.tgl_selesai, CUTTER.k_awal , CUTTER.k_akhir, " & _
                        "CUTTER.Keterangan, CUTTER.kondisiMesin, CUTTER.KondisiAgar " & _
                        " FROM CUTTER INNER JOIN " & _
                        " [Manufacture Order] ON CUTTER.OrderID = [Manufacture Order].OrderID " & _
                        " WHERE [Manufacture Order].Status <> 'RELEASED'"
   End If
   
   End With
   HiasFormManTell Picture2, Me
   Set mCall = New frmCaller
End Sub
Private Sub DDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
   cmdLink(0).Enabled = False
   cmdLink(1).Enabled = False
   PrepareSQL
   Select Case AdReasonActiveDb
      Case tmbAddNew:
         cmdLink(0).Enabled = True
         cmdLink(1).Enabled = True
         MEdit = False
      Case tmbEdit:
         cmdLink(0).Enabled = True
         cmdLink(1).Enabled = True
         MEdit = True
      Case tmbSave:
         DDE.IsChildMemberReady = True
      Case tmbDelete:
         
   End Select
End Sub

Private Sub PrepareSQL()
   DDE.PrepareAppend = "insert into CUTTER (no_ekstraksi,orderID, tanggal_cutter,grup,tgl_mulai,tgl_selesai,k_awal,k_akhir," & _
                                            "keterangan,KondisiMesin,KondisiAgar,issued_by) " & _
                       " values ('" & txt(0).Text & "', '" & txt(2).Text & "','" & _
                                 Format(tgl_cutter.Value, "yyyy-MM-dd hh:mm:ss") & "','" & _
                                 DDE.GetFieldByName("grup") & "', '" & _
                                 Format(tgl_awal.Value, "yyyy-MM-dd hh:mm:ss") & "','" & _
                                 Format(tgl_akhir.Value, "yyyy-MM-dd hh:mm:ss") & "','" & _
                                 DDE.GetFieldByName("k_awal") & "','" & DDE.GetFieldByName("k_akhir") & "','" & _
                                 DDE.GetFieldByName("keterangan") & "'," & _
                                 IIf(OptMesinA.Value = True, "1", "0") & "," & _
                                 IIf(OptAgarA.Value = True, "1", "0") & ",'" & MainMenu.StatusBar1.Panels(1) & "')"
   
   DDE.PrepareUpdate = "update cutter set OrderID = '" & txt(2).Text & "'," & _
                                        " tanggal_cutter = '" & Format(tgl_cutter.Value, "yyyy-MM-dd hh:mm:ss") & "'," & _
                                        " grup = '" & DDE.GetFieldByName("grup") & "'," & _
                                        " tgl_mulai = '" & Format(tgl_awal.Value, "yyyy-MM-dd hh:mm:ss") & "' ," & _
                                        " tgl_selesai = '" & Format(tgl_akhir.Value, "yyyy-MM-dd hh:mm:ss") & "'," & _
                                        " k_awal ='" & DDE.GetFieldByName("k_awal") & "'," & _
                                        " k_akhir ='" & DDE.GetFieldByName("k_akhir") & "'," & _
                                        " Keterangan = '" & DDE.GetFieldByName("keterangan") & "'," & _
                                        " KondisiMesin=" & IIf(OptMesinA.Value = True, "1", "0") & "," & _
                                        " KondisiAgar=" & IIf(OptAgarA.Value = True, "1", "0") & _
                        " where no_ekstraksi = '" & txt(0).Text & "'"
   
   DDE.PrepareDelete = "delete from CUTTER where no_ekstraksi = '" & DDE.GetFieldByName("no_ekstraksi") & "'"
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
   Select Case UCase(mCall.FromTagActive)
      Case "MANUFACTURE ORDER"
         txt(2).Text = mCall.GetFieldByName(0)
      Case "NO EKSTRAKSI"
         txt(0).Text = mCall.GetFieldByName(0)
   End Select
End Sub


Private Sub SaveToMO()
   Dim dStart As Date
   Dim dFinish As Date
   Dim ActualTime As Double
   Dim sWCID As String
   Dim rsCek As New DBQuick
   
   dStart = tgl_awal.Value
   dFinish = tgl_akhir.Value
   
   ActualTime = Val(SelisihHariJam(dStart, dFinish, 2))
   
   rsCek.DBOpen "select WCID from WCenter_Header where FormID = 44", CNN
   If rsCek.DBRecordset.Recordcount > 0 Then
      sWCID = rsCek.DBRecordset.Fields(0)
      SendDataToServer "update [order output detail] set actual_time=" & ActualTime & " where OrderID='" & DDE.GetFieldByName("OrderID") & "' and WCID='" & sWCID & "'"
   End If
   rsCek.CloseDB
End Sub

Private Sub tgl_akhir_Change()
   If tgl_awal.Value > tgl_akhir.Value Then
      MessageBox "Tgl & Waktu selesai tidak boleh lebih kecil dari tgl & Waktu Selesai ", "Peringatan", msgOkOnly, msgCrtical
      tgl_akhir.Value = tgl_awal.Value
   End If
End Sub

Private Sub tgl_awal_Change()
   If tgl_awal.Value > tgl_akhir.Value Then
      MessageBox "Tgl & Waktu selesai tidak boleh lebih kecil dari tgl & Waktu Selesai ", "Peringatan", msgOkOnly, msgCrtical
      tgl_awal.Value = tgl_akhir.Value
   End If

End Sub

Private Sub txt_LostFocus(Index As Integer)
   If Index = 0 Then
      Dim rsCek As New DBQuick
      rsCek.DBOpen "select * from statusProduksi where noEkstraksi='" & txt(0).Text & "'", CNN, lckLockBatch
      If rsCek.DBRecordset.Recordcount > 0 Then
         rsCek.DBOpen "select * from CUTTER where no_Ekstraksi='" & txt(0).Text & "'", CNN, lckLockBatch
         If rsCek.DBRecordset.Recordcount > 0 Then
            MessageBox "Nomor Ekstraksi Ini Sudah Diinput...!", "Peringatan", msgOkOnly, msgCrtical
            txt(0).Text = ""
         End If
      Else
         MessageBox "Nomor Ekstraksi Ini tidak ditemukan...!", "Peringatan", msgOkOnly, msgCrtical
         txt(0).Text = ""
      End If
      rsCek.CloseDB
   End If
End Sub
