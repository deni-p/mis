VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{D7BB8F75-AC9E-4E80-A526-70EA20ACFD16}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmProdAcidTreatment 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acid Treatment1"
   ClientHeight    =   6210
   ClientLeft      =   -45
   ClientTop       =   375
   ClientWidth     =   9105
   Icon            =   "FrmProdAcidTreatment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9105
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5535
      ScaleWidth      =   9105
      TabIndex        =   0
      Top             =   0
      Width           =   9105
      Begin VB.TextBox txtTanki 
         Appearance      =   0  'Flat
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5400
         TabIndex        =   15
         Tag             =   "ACID"
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   5400
         MultiLine       =   -1  'True
         TabIndex        =   7
         Tag             =   "ACID"
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txtGroup 
         Appearance      =   0  'Flat
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5400
         TabIndex        =   1
         Tag             =   "ACID"
         Top             =   120
         Width           =   2055
      End
      Begin MSFlexGridLib.MSFlexGrid GridAcid 
         Height          =   3945
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   8925
         _ExtentX        =   15743
         _ExtentY        =   6959
         _Version        =   393216
         Cols            =   8
         BackColor       =   16777215
         BackColorFixed  =   16761024
         GridColorFixed  =   16777215
         AllowBigSelection=   0   'False
         HighLight       =   0
         GridLinesFixed  =   1
         AllowUserResizing=   1
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DcTanggal 
         DataField       =   "DateTrans"
         Height          =   315
         Left            =   1440
         TabIndex        =   16
         Tag             =   "ACID"
         Top             =   840
         Width           =   2535
         _ExtentX        =   4471
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
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   58916867
         CurrentDate     =   39634
      End
      Begin VB.Label lblDokumentNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Ekstraksi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   165
         Width           =   1035
      End
      Begin VB.Label lblDokNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataSource      =   "DataTrans"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         Tag             =   "ACID"
         Top             =   120
         Width           =   1845
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   8
         X1              =   1430
         X2              =   120
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label lblEkstraksi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataSource      =   "DataTrans"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1440
         TabIndex        =   11
         Tag             =   "ACID"
         Top             =   480
         Width           =   2190
      End
      Begin VB.Label lblKeterangan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   4200
         TabIndex        =   10
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   5440
         X2              =   4200
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label lblNoEkstraksi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rekomendasi"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   525
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   5
         X1              =   5440
         X2              =   4200
         Y1              =   1365
         Y2              =   1365
      End
      Begin VB.Label lblKeterangan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanki"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   4200
         TabIndex        =   6
         Top             =   540
         Width           =   585
      End
      Begin VB.Label lblGroup 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   4200
         TabIndex        =   5
         Top             =   180
         Width           =   585
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   5440
         X2              =   4200
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label lblUOM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   900
         Width           =   585
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   1430
         X2              =   120
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label lblid 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000005&
         Height          =   210
         Left            =   14640
         TabIndex        =   2
         Top             =   7560
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   1430
         X2              =   120
         Y1              =   1140
         Y2              =   1140
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   4
      Top             =   5640
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   1005
      BindFormTAG     =   "ACID"
      ActiveLanguage  =   1
   End
   Begin VB.Label lblKeterangan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   8
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
Attribute VB_Name = "FrmProdAcidTreatment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdEkstraksi_Click()
    OpenPartner 0
End Sub

Private Function OpenPartner(ByVal Index As Integer) As Boolean
    Set mCall = New frmCaller
    
    Select Case Index

        Case 0
            RcProduksi.DBOpen "SELECT LabRekomEkstraksi.RLNO,LabRekomEkstraksi.SplNo From LabRekomEkstraksi", CNN, lckLockReadOnly

    End Select
    
    If RcProduksi.Recordcount <> 0 Then

        Select Case Index

            Case 0
                mCall.FromTagActive = "ACID"

        End Select

        Set mCall.FormData = RcProduksi.DBRecordset
        mCall.LookUp Me
    Else

        MessageBox "Konfigurasi ACID TREATMENT masih kosong", "Peringatan", msgOkOnly
        OpenPartner = True
    End If

End Function

Private Sub mCall_RowColChange(ByVal TagForm As String, _
                               ByVal pRecordset As ADODB.Recordset)

    Select Case TagForm

        Case "ACID"
            lblEkstraksi.Caption = mCall.GetFieldByName("RlNo")
    End Select

End Sub

Private Sub Form_Activate()
    lblEkstraksi = frmProduksi.txtBox(5)
    OpenDetail lblEkstraksi
    MergeGrid
End Sub

Private Sub OpenDetail(ByVal ParameterString As String)
    Dim nCount As Integer
    Set RcDetail = New DBQuick

    If ParameterString = "" Then ParameterString = "11111111111" ': Exit Sub
  
    RcDetail.DBOpen "SELECT LabRekomEkstraksi.SplNo,LabRekomEkstraksi.RLNO,LabRekomEkstraksi_Line.FORMID,LabRekomEkstraksi_Line.FORMNAME,LabProses.Prosedur, LabAnalysis.Analysis, LabSetupRekom_Line.minvalue, LabSetupRekom_Line.maxvalue From  LabRekomEkstraksi_Line INNER JOIN LabRekomEkstraksi ON (LabRekomEkstraksi_Line.SplNo = LabRekomEkstraksi.SplNo)  INNER JOIN LabSetupRekom_Header ON (LabRekomEkstraksi_Line.FORMID = LabSetupRekom_Header.FormID) INNER JOIN LabSetupRekom_Line ON (LabSetupRekom_Header.DocID = LabSetupRekom_Line.DocID)  AND (LabSetupRekom_Header.FormID = LabSetupRekom_Line.FormID) " & "INNER JOIN LabAnalysis ON (LabSetupRekom_Line.ID_ANALYSIS = LabAnalysis.ID) INNER JOIN LabProses ON (LabSetupRekom_Line.ProsesID = LabProses.ProsesID) Where  LabRekomEkstraksi.SplNo = '" & lblEkstraksi & "'  and LabRekomEkstraksi_Line.FORMNAME = 'ACID TREATMENT' Order By  LabSetupRekom_Line.ProsesID ", CNN, lckLockBatch
          
    Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
    GridAcid.Rows = 1

    With GridAcid

        For nCount = 0 To MyDDE.ChildRecordset.Recordcount - 1
            .AddItem ""
            .TextMatrix(nCount + 1, 1) = MyDDE.ChildRecordset.Fields("Prosedur")
            .TextMatrix(nCount + 1, 2) = MyDDE.ChildRecordset.Fields("Analysis")
            MyDDE.ChildRecordset.MoveNext
        Next nCount

        .MergeCol(1) = True
        .MergeCells = flexMergeFree
        .FixedCols = 2
        .RowHeight(0) = 500

    End With
   
    RcDetail.CloseDB
   
End Sub

Private Sub MergeGrid()
    Dim FixedColCaptions(1 To 6) As String, C As Long, I
    GridAltColor = &HEEDAC1
    
    Dim nCount As Integer
    Changingsel = 1
    GridAcid.FillStyle = flexFillRepeat
    GridAcid.Redraw = True

    With GridAcid
        .Cols = 4
        .ColWidth(0) = 350
        .TextMatrix(0, 1) = "PROSEDUR"
        .ColWidth(1) = 1800
        .TextMatrix(0, 2) = "ANALISA"
        .ColWidth(2) = 2200
        .TextMatrix(0, 3) = "KEBUTUHAN PRODUKSI"
        .ColWidth(3) = 1900
      
        For nCount = 1 To .Rows - 1
            .Row = nCount
            .Col = 1
            .RowSel = nCount
            .ColSel = .Cols - 1

            If .Rows > 1 Then
                If (.TextMatrix(nCount, 1) = .TextMatrix(nCount - 1, 1)) And .Row > 1 Then
                    If (.TextMatrix(nCount, 1) = .TextMatrix(nCount - 1, 1)) And (.TextMatrix(nCount, 1) = .TextMatrix(nCount - 1, 1)) Then
                        .CellBackColor = GridAltColor
                        Me.Tag = .TextMatrix(nCount, 1)
                    Else
                        .CellBackColor = &H8000000B
                    End If

                Else

                    If .Row = 1 Then
                        .CellBackColor = &H8000000B
                    ElseIf .TextMatrix(nCount, 1) = Me.Tag Then
                        .CellBackColor = GridAltColor
                        Me.Tag = .TextMatrix(nCount, 1)
                    Else
                        .CellBackColor = &H8000000B
                        Me.Tag = .TextMatrix(nCount, 1)
                        ' .TextMatrix(nCount, 1) = ""
                    End If
                End If

                '.CellBackColor = GridAltColor
            End If

        Next nCount

        .Gridlines = flexGridFlat
    End With

End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb

        Case adAddNew

            MEdit = True

        Case tmbSave

            If MyDDE.CheckEmptyControl = False Then
                MyDDE.IsChildMemberReady = True
                PrepareQuery
            Else
                MyDDE.IsChildMemberReady = False
            End If
            
        Case tmbAddNew
            MEdit = True
            Me.Tag = "Baru"
            lblDokNo.Caption = IndexAuto
            DcTanggal.Enabled = True
            DcTanggal.SetFocus
    End Select

End Sub

Private Sub PrepareQuery()
    On Error Resume Next

    With MyDDE

        .PrepareAppend = "INSERT INTO LabProsesProduksi_Header(DokNo, NoEkstraksi, RLNo, Berat, Tanggal,[Group], Tanki, Keterangan, Kondisi) VALUES ('" & lblDokNo.Caption & "','" & lblEkstraksi.Caption & "','" & lblNoRL & "','" & txtBerat.Text & "','" & DcTanggal.value & "','" & txtGroup.Text & "','" & txtTanki.Text & "','" & txtKeterangan.Text & "', sd)"
        .PrepareUpdate = "UPDATE LabRekomEkstraksi SET  RLNO = '" & lblDokNo.Caption & "', tempatalkali= '" & "1" & "' Where  LabRekomEkstraksi.SplNo ='" & .GetFieldByName("SplNo") & "'"
        .PrepareDelete = " DELETE FROM  [LabRekomEkstraksi] WHERE (SplNo = '" & .GetFieldByName("SplNo") & "')"
    End With

End Sub

Private Function IndexAuto() As String
    On Error Resume Next
    Dim Rc As New DBQuick
    Dim TglSaiki As String
    Dim Inom As String
    TglSaiki = Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2)
    Rc.DBOpen "SELECT  MAX(DokNo) AS MaxNom FROM [LabProsesProduksi_Header]", CNN, lckLockReadOnly

    With Rc

        If .DBRecordset.Recordcount <> 0 Then
            Inom = IIf(Not IsNull(.Fields(0)), Mid(.DBRecordset.Fields("MaxNom"), 12, 5), "0") + 1

            If Err.Number = 94 Then Inom = 1
        Else
            Inom = 1
        End If

        Select Case Len(Trim(Str(Inom)))

            Case 0
                IndexAuto = "PAC-" & TglSaiki & "-" & Trim(Str(Inom))

            Case 1
                IndexAuto = "PAC-" & TglSaiki & "-" & "0000" & Trim(Str(Inom))

            Case 2
                IndexAuto = "PAC-" & TglSaiki & "-" & "000" & Trim(Str(Inom))

            Case 3
                IndexAuto = "PAC-" & TglSaiki & "-" & "00" & Trim(Str(Inom))

            Case 4
                IndexAuto = "PAC-" & TglSaiki & "-" & "0" & Trim(Str(Inom))
        End Select

    End With

End Function
