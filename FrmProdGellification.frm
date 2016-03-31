VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmProdGellification 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gellification"
   ClientHeight    =   4545
   ClientLeft      =   225
   ClientTop       =   540
   ClientWidth     =   9420
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmProdGellification.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4545
   ScaleWidth      =   9420
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   0
      ScaleHeight     =   3975
      ScaleWidth      =   9420
      TabIndex        =   10
      Top             =   0
      Width           =   9420
      Begin VB.TextBox txtGroup 
         Appearance      =   0  'Flat
         DataField       =   "Group"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Tag             =   "GEL"
         Top             =   1215
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DcTanggal 
         DataField       =   "Tanggal"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Tag             =   "GEL"
         Top             =   840
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   70909955
         CurrentDate     =   39634
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "tanggal_mulai"
         DataSource      =   "dde"
         Height          =   315
         Index           =   0
         Left            =   6045
         TabIndex        =   7
         Tag             =   "dryer"
         Top             =   840
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMM yyyy    hh:mm"
         Format          =   70909955
         CurrentDate     =   39419
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "tanggal_selesai"
         DataSource      =   "dde"
         Height          =   315
         Index           =   1
         Left            =   6060
         TabIndex        =   8
         Tag             =   "dryer"
         Top             =   1200
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMM yyyy    hh:mm"
         Format          =   70909955
         CurrentDate     =   39419
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         DataField       =   "Keterangan"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   5160
         MultiLine       =   -1  'True
         TabIndex        =   5
         Tag             =   "GEL"
         Top             =   120
         Width           =   3015
      End
      Begin VB.CommandButton cmdRefLink 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7005
         Picture         =   "FrmProdGellification.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "BAHAN"
         Top             =   480
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton cmdEkstraksi 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3405
         Picture         =   "FrmProdGellification.frx":6BDC
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "Formula"
         Top             =   480
         Visible         =   0   'False
         Width           =   345
      End
      Begin MSFlexGridLib.MSFlexGrid GridFilterPress 
         Height          =   2145
         Left            =   120
         TabIndex        =   9
         Top             =   1665
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   3784
         _Version        =   393216
         Cols            =   8
         BackColor       =   16777215
         GridColorFixed  =   16777215
         AllowBigSelection=   0   'False
         HighLight       =   0
         GridLinesFixed  =   1
         AllowUserResizing=   1
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblDokNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "DokNo"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Tag             =   "GEL"
         Top             =   120
         Width           =   1845
      End
      Begin VB.Label lblEksNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "NoEkstraksi"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Tag             =   "GEL"
         Top             =   480
         Width           =   2085
      End
      Begin VB.Label LbRefID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "refid"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5160
         TabIndex        =   20
         Tag             =   "GEL"
         Top             =   465
         Width           =   1845
      End
      Begin VB.Label lblTanggalWaktu 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal && waktu selesai"
         Height          =   255
         Index           =   4
         Left            =   3960
         TabIndex        =   22
         Top             =   1230
         Width           =   2055
      End
      Begin VB.Label lblTanggalWaktu 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal && waktu mulai"
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   21
         Top             =   870
         Width           =   1890
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   6
         X1              =   6040
         X2              =   3960
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   7
         X1              =   6040
         X2              =   3960
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   13
         X1              =   5200
         X2              =   3960
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label lblReference 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3960
         TabIndex        =   19
         Top             =   525
         Width           =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   10
         X1              =   6160
         X2              =   4920
         Y1              =   3495
         Y2              =   3495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   8
         X1              =   1360
         X2              =   120
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label lblNoDokumen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Dokumen"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   165
         Width           =   1125
      End
      Begin VB.Label lblKeterangan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   3960
         TabIndex        =   16
         Top             =   180
         Width           =   945
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   5
         X1              =   5200
         X2              =   3960
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label lblGroup 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   1275
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   1360
         X2              =   120
         Y1              =   1515
         Y2              =   1515
      End
      Begin VB.Label lblTanggal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   900
         Width           =   645
      End
      Begin VB.Label lblid 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Left            =   14640
         TabIndex        =   13
         Top             =   7560
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   1360
         X2              =   120
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Label lblNoEkstraksi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Ekstraksi"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   555
         Width           =   1035
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   1360
         X2              =   120
         Y1              =   825
         Y2              =   825
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   3975
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   1005
      BindFormTAG     =   "GEL"
      ActiveLanguage  =   1
   End
   Begin VB.Label lblKeterangan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   585
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   1240
      X2              =   0
      Y1              =   255
      Y2              =   255
   End
End
Attribute VB_Name = "FrmProdGellification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private RcProses As New DBQuick

Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1

Private RcDetail As New DBQuick

Private RsDetail As DBQuick

Private rcBarCode As DBQuick

Private MEdit As Boolean

Private rsCombo As DBQuick

Private rsPriority As DBQuick

Private rsLab As DBQuick
Dim IDGen As New IDGenerator
Dim movComplit As Boolean
Dim Xval As String
Dim GridAltColor As String
Dim Changingsel As Byte
Dim mFirstCaller As Boolean
Dim strSQL As String
Dim sWCID As String

Private Sub BindToGrid()

    If movComplit = False Then
        OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("DokNo")), MyDDE.GetFieldByName("DokNo"), "xxx")
        movComplit = False

    Else
        OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("DokNo")), MyDDE.GetFieldByName("DokNo"), "xxx")
        movComplit = False
    End If

End Sub

Private Function CekGridKosong() As Boolean
    Dim nRow As Integer

    For nRow = 1 To GridFilterPress.Rows - 1
        GridFilterPress.row = nRow

        If GridFilterPress.TextMatrix(nRow, 3) = "" Then
            CekGridKosong = True
            MessageBox "Baris ke-" & nRow + 1 & " Pada Kolom Hasil Pemeriksaan Harus Diisi", "Peringatan"
            Exit For
        End If

    Next nRow

End Function

'Private Sub DetailBarang(ByVal ParameterString As String)
'  Set RcDetail = New DBQuick
'  RcDetail.DBOpen _
'          "SELECT  [PO Order].PurchaseID,Inventory.ItemName,PartnerDB.CompanyName, PartnerDB.Address, PartnerDB.City From [PO Order] INNER JOIN PartnerDB ON ([PO Order].PartnerID = PartnerDB.PartnerID) INNER JOIN [Detail PO] ON ([PO Order].PurchaseID = [Detail PO].PurchaseID) INNER JOIN Inventory ON ([Detail PO].NoItem = Inventory.NoItem) Where  ([PO Order].StatusSJ = 0) AND  (LEFT([PO Order].PurchaseID, 2) = 'PO') AND ([PO Order].PurchaseID = '" _
'          & MyDDE.GetFieldByName("SPPHID") & "') Order By  [PO Order].PurchaseID ", CNN, lckLockBatch
'  txtLot.Caption = RcDetail.Fields("ItemName")
'  LblMaster(2) = RcDetail.Fields("CompanyName")
'End Sub

Private Sub DetailProsesID(ByVal ParameterString As String)
    Set RcDetail = New DBQuick
    RcDetail.DBOpen "SELECT ProdFormulaEkstraksi.EksNo From ProdFormulaEkstraksi Where (ProdFormulaEkstraksi.typeTrans = '" & ParameterString & "')", CNN, lckLockBatch
End Sub

Private Sub detilGrid(ByVal ParameterString As String)
    Set RcDetail = New DBQuick

    If ParameterString = "" Then ParameterString = "11111111111" ': Exit Sub
    RcDetail.DBOpen _
       "SELECT ProdProsesProduksi_Header.DokNo, LabSample_RL_Line.ProsesID, LabSample_RL_Line.TAT, LabProses.Prosedur, LabSample_RL_Line.Analysis, LabSample_RL_Line.Result, labconfigproses.MinValue, labconfigproses.MaxValue," _
       & _
       " LabAnalysis.unit, PartnerDB.CompanyName From  LabSample_RL_Line INNER JOIN ProdProsesProduksi_Header ON (LabSample_RL_Line.DokNo = ProdProsesProduksi_Header.DokNo) INNER JOIN ProdFormulaEkstraksi ON (ProdProsesProduksi_Header.EksNo = ProdFormulaEkstraksi.EksNo) INNER JOIN ProdFormulaEkstraksi_Detail ON (ProdFormulaEkstraksi.EksNo = ProdFormulaEkstraksi_Detail.EksNo) INNER JOIN LabProses ON (ProdFormulaEkstraksi_Detail.ProsesID = LabProses.ProsesID) INNER JOIN labconfigproses ON (ProdFormulaEkstraksi_Detail.ProsesID = labconfigproses.ProsesID)  AND (labconfigproses.ProsesID = LabProses.ProsesID)  AND (LabSample_RL_Line.Analysis = labconfigproses.Analysis)  INNER JOIN LabAnalysis ON (labconfigproses.Analysis = LabAnalysis.Analysis)  INNER JOIN [PO Order] ON (ProdProsesProduksi_Header.SPPHID = [PO Order].PurchaseID) " _
       & "INNER JOIN PartnerDB ON ([PO Order].PartnerID = PartnerDB.PartnerID) Where  ProdProsesProduksi_Header.DokNo = '" & _
       MyDDE.GetFieldByName("DokNo") & "'", CNN, lckLockBatch
       
    Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
    
    RcDetail.CloseDB
    isiGrid
End Sub

Private Sub SimpanDetail()
    Dim nRow As Integer

    With GridFilterPress

        If SendDataToServer("DELETE FROM [ProdProsesProduksi_Line] WHERE     (DokNo = N'" & lblDokNo.Caption & "')") = True Then

            For nRow = 0 To GridFilterPress.Rows - 2

                GridFilterPress.row = nRow
                SendDataToServer "INSERT INTO  ProdProsesProduksi_Line(DokNo, NoEkstraksi, ProsesID,ID,Analysis,Result) VALUES('" & lblDokNo.Caption & "','" & lbleksno & "','" & .TextMatrix(nRow + 1, 5) & "','" & .TextMatrix(nRow + 1, 6) & "','" & .TextMatrix(nRow + 1, 2) & "','" & .TextMatrix(nRow + 1, 3) & "')"
            
            Next nRow

            'SendDataToServer "UPDATE StatusProduksi SET Rekomendasi = '" & lblDokNo.Caption & "',Posisi = '" & "GELLIFICATION" & "', status = '1',tanggal = '" & DcTanggal.Value & "' Where  StatusProduksi.NoEkstraksi = '" & lbleksno.Caption & "'"
        End If

    End With

End Sub

Private Sub GridFilterPressNya()
    Dim FixedColCaptions(1 To 6) As String, C As Long, I
    GridAltColor = &HEEDAC1
    
    Dim ncount As Integer
    Changingsel = 1
    GridFilterPress.FillStyle = flexFillRepeat
    GridFilterPress.Redraw = True

    With GridFilterPress
        .Cols = 7
        .ColWidth(0) = 350
        .TextMatrix(0, 1) = "PROSEDUR"
        .ColWidth(1) = 1800
        .TextMatrix(0, 2) = "ANALISA"
        .ColWidth(2) = 2200
    
        .TextMatrix(0, 3) = "HASIL PEMERIKSAAN"
        .ColWidth(3) = 1800
    
        .TextMatrix(0, 4) = "SATUAN"
        .ColWidth(4) = 780
    
        .TextMatrix(0, 5) = "PROSES ID"
        .ColWidth(5) = 0
        .ColWidth(6) = 0
    
        For ncount = 1 To .Rows - 1
            .row = ncount
            .col = 1
            .RowSel = ncount
            .ColSel = .Cols - 1

            If .Rows > 1 Then
                If (.TextMatrix(ncount, 1) = .TextMatrix(ncount - 1, 1)) And .row > 1 Then
                    If (.TextMatrix(ncount, 1) = .TextMatrix(ncount - 1, 1)) And (.TextMatrix(ncount, 1) = .TextMatrix(ncount - 1, 1)) Then
                        .CellBackColor = GridAltColor
                        Me.Tag = .TextMatrix(ncount, 1)
                    Else
                        .CellBackColor = &H8000000B
                    End If

                Else

                    If .row = 1 Then
                        .CellBackColor = &H8000000B
                    ElseIf .TextMatrix(ncount, 1) = Me.Tag Then
                        .CellBackColor = GridAltColor
                        Me.Tag = .TextMatrix(ncount, 1)
                    Else
                        .CellBackColor = &H8000000B
                        Me.Tag = .TextMatrix(ncount, 1)
                        ' .TextMatrix(nCount, 1) = ""
                    End If
                End If

                '.CellBackColor = GridAltColor
            End If

        Next ncount

        .GridLines = flexGridFlat
    End With

End Sub

Private Function IndexAuto() As String
    On Error Resume Next
    Dim Rc As New DBQuick
    Dim TglSaiki As String
    Dim Inom As String
    TglSaiki = Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2)
    Rc.DBOpen "SELECT  MAX(DokNo) AS MaxNom FROM [ProdProsesProduksi_Header] where left(dokno,3)='GEL'", CNN, lckLockReadOnly

    With Rc

        If .DBRecordset.Recordcount <> 0 Then
            Inom = IIf(Not IsNull(.Fields(0)), Mid(.DBRecordset.Fields("MaxNom"), 12, 5), "0") + 1

            If Err.Number = 94 Then Inom = 1
        Else
            Inom = 1
        End If

        Select Case Len(Trim(Str(Inom)))

            Case 0
                IndexAuto = "GEL-" & TglSaiki & "-" & Trim(Str(Inom))

            Case 1
                IndexAuto = "GEL-" & TglSaiki & "-" & "0000" & Trim(Str(Inom))

            Case 2
                IndexAuto = "GEL-" & TglSaiki & "-" & "000" & Trim(Str(Inom))

            Case 3
                IndexAuto = "GEL-" & TglSaiki & "-" & "00" & Trim(Str(Inom))

            Case 4
                IndexAuto = "GEL-" & TglSaiki & "-" & "0" & Trim(Str(Inom))
        End Select

    End With

End Function

Private Sub isiGrid()
    Dim nRow As Byte
    Set RcDetail = New DBQuick
    RcDetail.DBOpen "SELECT  [PO Order].PurchaseID,Inventory.ItemName,PartnerDB.CompanyName, PartnerDB.Address, PartnerDB.City From [PO Order] INNER JOIN PartnerDB ON ([PO Order].PartnerID = PartnerDB.PartnerID) INNER JOIN [Detail PO] ON ([PO Order].PurchaseID = [Detail PO].PurchaseID) INNER JOIN Inventory ON ([Detail PO].NoItem = Inventory.NoItem) Where  ([PO Order].StatusSJ = 0) AND  (LEFT([PO Order].PurchaseID, 2) = 'PO') AND ([PO Order].PurchaseID = '" & MyDDE.GetFieldByName("SPPHID") & "') Order By  [PO Order].PurchaseID ", CNN, lckLockBatch
End Sub

'Private Function isNullGrid() As Boolean
'  Dim nRow As Byte
'
'  If MyDDE.ChildRecordset.Recordcount < 1 Then Exit Function
'
'  For nRow = 0 To MyDDE.ChildRecordset.Recordcount - 1
'    GrdRLLuar.Row = nRow
'
'    If GrdRLLuar.Columns(2).Text = "" Then MessageBox "Hasil Pemeriksaan Harus Diisi", "Peringatan"
'  Next nRow
'
'End Function

Private Sub loadAwal()

    With MyDDE

        If .ChildRecordset.Recordcount < 1 Then Exit Sub
        lblDokNo = IIf(IsNull(MyDDE.GetFieldByName("DokNo")) Or MyDDE.GetFieldByName("DokNo") = "", "", MyDDE.GetFieldByName("DokNo"))

        If MyDDE.ActiveRecordset.EOF Or .ActiveRecordset.BOF Then Exit Sub
    End With

End Sub

'Private Sub loadDetail()
'
'  With MyDDE.ActiveRecordset
'    ' If .Recordcount <> 0 Then
'    Set rsDetail = New DBQuick
'    strSQL = _
'            "SELECT DISTINCT ProdFormulaEkstraksi.EksNo, LabProses.Prosedur, labconfigproses.Analysis, labconfigproses.Methods, labconfigproses.MinValue, labconfigproses.MaxValue, LabAnalysis.unit, LabSample_RL_Line.Result " _
'            & _
'            " From ProdFormulaEkstraksi_Detail INNER JOIN ProdFormulaEkstraksi ON (ProdFormulaEkstraksi_Detail.EksNo = ProdFormulaEkstraksi.EksNo) INNER JOIN labconfigproses ON (ProdFormulaEkstraksi_Detail.ProsesID = labconfigproses.ProsesID) INNER JOIN LabProses ON (labconfigproses.ProsesID = LabProses.ProsesID) INNER JOIN LabAnalysis ON (labconfigproses.Analysis = LabAnalysis.Analysis) INNER JOIN LabSample_RL_Line ON (labconfigproses.ProsesID = LabSample_RL_Line.ProsesID) AND (LabSample_RL_Line.ProsesID = LabProses.ProsesID) AND (labconfigproses.Analysis = LabSample_RL_Line.Analysis) Where  ProdFormulaEkstraksi.typeTrans = 'SPL-RLUAR'"

'
'    rsDetail.DBOpen strSQL, CNN
'    Set MyDDE.ChildRecordset = rsDetail.DBRecordset.Clone(adLockBatchOptimistic)
'    Set GrdRLLuar.DataSource = MyDDE.ChildRecordset
'    '  End If
'  End With
'
'End Sub
'
Private Sub OpenDetail(ByVal ParameterString As String)
    Dim ncount As Integer
    Set RcDetail = New DBQuick

    If ParameterString = "" Then ParameterString = "11111111111" ': Exit Sub
     
    RcDetail.DBOpen "SELECT  ProdProsesProduksi_Header.DokNo,ProdProsesProduksi_Header.NoEkstraksi,ProdProsesProduksi_Header.Tanggal, ProdAnalysis.unit,ProdProsesProduksi_Header.[Group],ProdProsesProduksi_Header.Keterangan,ProdProsesProduksi_Line.ProsesID,ProdProses.Prosedur,ProdProsesProduksi_Line.Analysis,ProdProsesProduksi_Line.Result From ProdProsesProduksi_Line INNER JOIN ProdProsesProduksi_Header ON (ProdProsesProduksi_Line.DokNo = ProdProsesProduksi_Header.DokNo) " & " INNER JOIN ProdProses ON (ProdProsesProduksi_Line.ProsesID = ProdProses.ProsesID) INNER JOIN ProdAnalysis ON (ProdProsesProduksi_Line.ID = ProdAnalysis.ID) Where ProdProsesProduksi_Header.DokNo = '" & MyDDE.GetFieldByName("DokNo") & "' order by ProdProsesProduksi_Line.ProsesID", CNN, lckLockBatch
    Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
    GridFilterPress.Rows = 1

    With GridFilterPress

        For ncount = 0 To MyDDE.ChildRecordset.Recordcount - 1
            .AddItem ""
            .TextMatrix(ncount + 1, 1) = MyDDE.ChildRecordset.Fields("Prosedur")
            .TextMatrix(ncount + 1, 2) = MyDDE.ChildRecordset.Fields("Analysis")
            .TextMatrix(ncount + 1, 4) = IIf(IsNull(MyDDE.ChildRecordset.Fields("unit")), "", MyDDE.ChildRecordset.Fields("unit"))
            .TextMatrix(ncount + 1, 3) = MyDDE.ChildRecordset.Fields("Result")
            .TextMatrix(ncount + 1, 5) = MyDDE.ChildRecordset.Fields("ProsesID")
            .TextMatrix(ncount + 1, 6) = MyDDE.ChildRecordset.Fields("ProsesID")
            MyDDE.ChildRecordset.MoveNext
        Next ncount

        .MergeCol(1) = True
        .MergeCells = flexMergeFree
        .FixedCols = 2
        .RowHeight(0) = 500

    End With

    RcDetail.CloseDB
    isiGrid
End Sub

Private Sub FormulaRL()
    Dim ncount As Integer

    With MyDDE.ActiveRecordset
        Set RsDetail = New DBQuick
        strSQL = "SELECT DISTINCT ProdAnalysis.id,ProdFormulaEkstraksi.EksNo,Prodconfigproses.ProsesID,ProdProses.Prosedur, Prodconfigproses.Analysis,Prodconfigproses.MinValue,Prodconfigproses.Methods," & " ProdAnalysis.unit,ProdFormulaEkstraksi.TypeTrans From ProdFormulaEkstraksi_Detail INNER JOIN ProdFormulaEkstraksi ON (ProdFormulaEkstraksi_Detail.EksNo =  ProdFormulaEkstraksi.EksNo) INNER JOIN Prodconfigproses ON (ProdFormulaEkstraksi_Detail.ProsesID =Prodconfigproses.ProsesID) " & "INNER JOIN ProdProses ON (Prodconfigproses.ProsesID = ProdProses.ProsesID) INNER JOIN ProdAnalysis ON (Prodconfigproses.Analysis = ProdAnalysis.Analysis) where ProdFormulaEkstraksi.typeTrans = 'PRO-GEL'"
        RsDetail.DBOpen strSQL, CNN

        If RsDetail.Recordcount < 1 Then
            MessageBox "Konfigurasi Form Gellification Masih Kosong.", "Peringatan", msgOkOnly
            cmdEkstraksi.Enabled = False
        End If

        Set MyDDE.ChildRecordset = RsDetail.DBRecordset.Clone(adLockBatchOptimistic)
      
        GridFilterPress.Rows = 1

        For ncount = 0 To MyDDE.ChildRecordset.Recordcount - 1
            GridFilterPress.AddItem ""
            GridFilterPress.TextMatrix(ncount + 1, 1) = MyDDE.ChildRecordset.Fields("Prosedur")
            GridFilterPress.TextMatrix(ncount + 1, 2) = MyDDE.ChildRecordset.Fields("Analysis")
            GridFilterPress.TextMatrix(ncount + 1, 4) = IIf(IsNull(MyDDE.ChildRecordset.Fields("unit")), "", MyDDE.ChildRecordset.Fields("unit"))
            GridFilterPress.TextMatrix(ncount + 1, 5) = MyDDE.ChildRecordset.Fields("ProsesID")
            GridFilterPress.TextMatrix(ncount + 1, 6) = MyDDE.ChildRecordset.Fields("ID")
            MyDDE.ChildRecordset.MoveNext
        Next ncount
   
    End With

    GridFilterPressNya
End Sub

Private Function OpenPartner(ByVal Index As Integer) As Boolean
    Set mCall = New frmCaller
    
    Select Case Index

        Case 0
            'RcProses.DBOpen "SELECT ProdProsesProduksi_Header.DokNo as [Nomor Ekstraksi] From ProdProsesProduksi_Header WHERE left (ProdProsesProduksi_Header.DokNo, 3) = 'FIL'", CNN, lckLockReadOnly
            RcProses.DBOpen "SELECT StatusProduksi.NoEkstraksi, StatusProduksi.Rekomendasi, StatusProduksi.Posisi, StatusProduksi.status, StatusProduksi.tanggal From StatusProduksi Where left(StatusProduksi.Rekomendasi,3) = 'FIL' AND StatusProduksi.status = 1 and StatusProduksi.Posisi='FILTERPRESS'", CNN, lckLockReadOnly

        Case 1
            RcProses.DBOpen "SELECT [Manufacture Order].OrderID From [Manufacture Order] Order By  [Manufacture Order].OrderID ", CNN, lckLockReadOnly
    End Select
    
    If RcProses.Recordcount <> 0 Then

        Select Case Index

            Case 0
                mCall.FromTagActive = "NOMOR EKSTRAKSI"

            Case 1
                mCall.FromTagActive = "REFERENCE"
        End Select

        Set mCall.FormData = RcProses.DBRecordset
        mCall.LookUp Me
    Else

        If Index = 0 Then MessageBox "Nomor Ekstraksi Masih Kosong. ", "Peringatan", msgOkOnly
        OpenPartner = True
        Exit Function
    End If

End Function

Private Sub PrepareQuery()
    On Error Resume Next
    Dim mPoSc As String
    Dim strSQL As String

    With MyDDE
        .PrepareAppend = "INSERT INTO ProdProsesProduksi_Header(DokNo,NoEkstraksi,Tanggal,[Group],Keterangan, refid,type_proses) VALUES('" & lblDokNo.Caption & "','" & lbleksno.Caption & "',CONVERT(DATETIME,'" & Format(DcTanggal.Value, "dd/mm/yy") & "',3),'" & txtGroup.Text & "','" & txtKeterangan.Text & "','" & LbRefID.Caption & "','" & "GELL" & "')"
     
        strSQL = "UPDATE ProdProsesProduksi_Header SET DokNo = '" & lblDokNo.Caption & "',refid='" & LbRefID.Caption & "',NoEkstraksi = '" & lbleksno.Caption & "',Tanggal =CONVERT(DATETIME,'" & Format(DcTanggal.Value, "dd/mm/yy") & "',3),[Group] = '" & txtGroup.Text & "', Keterangan = '" & txtKeterangan.Text & "' Where ProdProsesProduksi_Header.DokNo = '" & .GetFieldByName("DokNo") & "'"

        .PrepareUpdate = strSQL
                     
        .PrepareDelete = " DELETE FROM  [ProdProsesProduksi_Header] WHERE (DokNo = '" & .GetFieldByName("DokNo") & "')"
    End With

    Err.Clear
End Sub

Private Sub cmdKP_Click()
    OpenPartner 1
End Sub

Private Sub cmdSPPH_Click()
    OpenPartner 0
End Sub

Private Sub cmdEkstraksi_Click()
    OpenPartner 0
End Sub

Private Sub cmdRefLink_Click()
    OpenPartner 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    ScanKey KeyCode, Shift, MyDDE

    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    lbleksno = frmProduksi.txtBox(5)
    LbRefID = frmProduksi.txtBox(0)
    movComplit = True
    HiasFormManTell Picture2, Me
  
    With MyDDE
        .EditModeReplace = False
        Set .BindForm = Me
        .BindFormTAG = "GEL"
        .SetPermissions = UserDeleteDenied
        Set .ActiveConnection = CNN
        .PrepareQuery = "select * from ProdProsesProduksi_Header where type_proses='GELL'"
        .SetPermissions = aksess.MayDo("Gellification")
    End With

    Set mCall = New frmCaller
    OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("DokNo")), MyDDE.GetFieldByName("DokNo"), "xxx")
    'GridFilterPressNya
    Me.Tag = "lock"
End Sub

Private Sub GridFilterPress_KeyDown(KeyCode As Integer, _
                                    Shift As Integer)

    If MEdit = False Then Exit Sub
    Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridFilterPress_KeyPress(KeyAscii As Integer)

    If (LCase(GridFilterPress.Tag) = "baru") And (GridFilterPress.col = 3) Then 'Or GridFilterPress.ColWidth(3) > 0 Then
        If KeyAscii = vbKeyReturn Then
            If GridFilterPress.col + 1 = GridFilterPress.Cols Then
                If GridFilterPress.row + 1 = GridFilterPress.Rows Then
                    GridFilterPress.row = 0
                    GridFilterPress.col = 0
                End If

                GridFilterPress.row = GridFilterPress.row + 1
                GridFilterPress.col = 0
            Else
                GridFilterPress.col = GridFilterPress.col + 1
            End If
        End If
    
        If KeyAscii = 8 Then
            If Len(Xval) = 0 Then Exit Sub
            Xval = Left$(Xval, Len(Xval) - 1)
            Exit Sub
        End If

        Xval = Xval & Chr(KeyAscii)
    End If
   
End Sub

Private Sub GridFilterPress_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If (GridFilterPress.col = 3) And (LCase(GridFilterPress.Tag) = "baru") Then
        GridFilterPress.Text = Xval
    End If

End Sub

Private Sub GridFilterPress_RowColChange()
    Xval = ""
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, _
                               ByVal pRecordset As ADODB.Recordset)

    Select Case TagForm

        Case "NOMOR EKSTRAKSI"

            With MyDDE
                lbleksno.Caption = mCall.GetFieldByName("NoEkstraksi")
            End With

        Case "REFERENCE"

            LbRefID.Caption = mCall.GetFieldByName("OrderID")
    End Select

End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
    Dim IDGen As New IDGenerator

    Select Case AdReasonActiveDb

        Case tmbDetail

            If mFirstCaller = False Then
                OpenPartner 1
                MEdit = True
            End If

        Case tmbAddNew

            MEdit = True
            Me.Tag = "baru"
            lblDokNo.Caption = IDGen.GetID("GELL")
            cmdEkstraksi.Enabled = MEdit
            
            lbleksno = frmProduksi.txtBox(5)
            LbRefID = frmProduksi.txtBox(0)
            txtGroup.SetFocus

            FormulaRL

        Case tmbSave

            If MyDDE.IsChildMemberReady = True Then
                If CekGridKosong = False Then
                    SimpanDetail
                   ' SaveToMO
                    MEdit = False
                End If
            End If

        Case tmbEdit
            MEdit = True

        Case tmbDelete
            PrepareQuery
    End Select

End Sub

Private Sub SaveToMO()
    Dim dStart As Date
    Dim dFinish As Date
    Dim ActualTime As Double
    Dim rsCek As New DBQuick
   
    dStart = tgl(0).Value
    dFinish = tgl(1).Value
    ActualTime = Val(SelisihHariJam(dStart, dFinish, 2))
   
    rsCek.DBOpen "select WCID from WCenter_Header where FormID = 42", CNN
    GetWC LbRefID.Caption

    If rsCek.DBRecordset.Recordcount > 0 Then
        sWCID = rsCek.DBRecordset.Fields(0)
        SendDataToServer "update [order output detail] set actual_time=" & ActualTime & " where OrderID='" & LbRefID.Caption & "' and WCID='" & sWCID & "'"
    End If

    rsCek.CloseDB
End Sub

Private Function GetWC(ByVal FormIDNya As String)
    Dim RcGetWC As New DBQuick
    RcGetWC.DBOpen "SELECT wcenter_header.WCID From wcenter_header Where wcenter_header.formid = 38", CNN, lckLockReadOnly
    sWCID = RcGetWC.DBRecordset.Fields("WCID")
End Function

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
                               ByVal pError As ADODB.Error, _
                               adStatus As ADODB.EventStatusEnum, _
                               ByVal pRecordset As ADODB.Recordset)
    OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("DokNo")), MyDDE.GetFieldByName("DokNo"), "xxx")
    GridFilterPressNya
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb

        Case tmbDelete
            PrepareQuery

        Case tmbAddNew

            MEdit = True
            GridFilterPress.Tag = "baru"
            lbleksno = frmProduksi.txtBox(5)
            LbRefID = frmProduksi.lblSplNo

        Case tmbSave

            If MyDDE.CheckEmptyControl = False Then
                If CekGridKosong = False Then  'And MyDDE.ChildRecordset.Recordcount <> 0 Then
                    MyDDE.IsChildMemberReady = True
                    PrepareQuery
                Else
                    MyDDE.IsChildMemberReady = False
                End If

            Else
                MyDDE.IsChildMemberReady = False
            End If
            
        Case 1
            MEdit = True

        Case 8, 9, 10, 11

            movComplit = True

        Case tmbDetail

            If MyDDE.CheckEmptyControl = False Then
            Else
                MyDDE.CancelTrans = mFirstCaller
            End If

        Case tmbEdit

            MEdit = True
         
            Me.Tag = "unlock"

        Case tmbCancel
            MEdit = False
            
    End Select

End Sub

