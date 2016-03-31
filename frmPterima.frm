VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{D7BB8F75-AC9E-4E80-A526-70EA20ACFD16}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmPterima 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SERAH TERIMA PRODUK JADI"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8910
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   8910
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3660
      Left            =   120
      ScaleHeight     =   3630
      ScaleWidth      =   8730
      TabIndex        =   0
      Top             =   0
      Width           =   8760
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   3525
         Left            =   60
         ScaleHeight     =   3495
         ScaleWidth      =   8595
         TabIndex        =   1
         Top             =   60
         Width           =   8625
         Begin MSDataGridLib.DataGrid DGDETAIL 
            Height          =   2175
            Left            =   120
            TabIndex        =   7
            Tag             =   "terima"
            Top             =   1200
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   3836
            _Version        =   393216
            AllowUpdate     =   -1  'True
            AllowArrows     =   -1  'True
            Appearance      =   0
            DefColWidth     =   6667
            HeadLines       =   1
            RowHeight       =   15
            WrapCellPointer =   -1  'True
            RowDividerStyle =   1
            FormatLocked    =   -1  'True
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "DAFTAR PRODUK"
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "lot_no"
               Caption         =   "LOT NO"
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
               DataField       =   "jumlah"
               Caption         =   "JUMLAH"
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
               DataField       =   "satuan"
               Caption         =   "SATUAN"
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
                  DividerStyle    =   4
                  Button          =   -1  'True
                  WrapText        =   -1  'True
                  ColumnWidth     =   1514,835
               EndProperty
               BeginProperty Column01 
               EndProperty
               BeginProperty Column02 
               EndProperty
            EndProperty
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "no_terima"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   1
            Left            =   2055
            TabIndex        =   2
            Tag             =   "terima"
            Top             =   435
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker tgl 
            DataField       =   "tanggal_ekstrasi"
            DataSource      =   "DDE"
            Height          =   315
            Left            =   2040
            TabIndex        =   3
            Tag             =   "terima"
            Top             =   720
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy"
            Format          =   58195971
            CurrentDate     =   39365
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   2160
            X2              =   120
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   2160
            X2              =   120
            Y1              =   735
            Y2              =   735
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal"
            Height          =   255
            Index           =   2
            Left            =   135
            TabIndex        =   5
            Top             =   810
            Width           =   2055
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Top             =   495
            Width           =   2055
         End
      End
   End
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   6
      Top             =   3795
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   1005
      BindFormTAG     =   "cruz"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmPterima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents lot As frmCaller
Attribute lot.VB_VarHelpID = -1
Dim rslot As New DBQuick

Private Function IndexAuto() As String
    Dim Rc As New DBQuick
    Dim TglSaiki As String
    Dim Inom As Long
    TglSaiki = Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2)
    Rc.DBOpen "SELECT MAX(RIGHT(No_terima, 5)) AS MaxNom FROM [t_produk_header] WHERE (GETDATE() = { fn NOW() })", CNN, lckLockReadOnly

    With Rc

        If .DBRecordset.Recordcount <> 0 Then
            Inom = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + 1
        Else
            Inom = 1
        End If

        Select Case Len(Trim(Str(Inom)))

            Case 0: IndexAuto = "ST/" & TglSaiki & "-" & Trim(Str(Inom))

            Case 1: IndexAuto = "ST/" & TglSaiki & "-" & "0000" & Trim(Str(Inom))

            Case 2: IndexAuto = "ST/" & TglSaiki & "-" & "000" & Trim(Str(Inom))

            Case 3: IndexAuto = "ST/" & TglSaiki & "-" & "00" & Trim(Str(Inom))

            Case 4: IndexAuto = "ST/" & TglSaiki & "-" & "0" & Trim(Str(Inom))
        End Select

    End With

End Function

Private Sub DDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb

        Case tmbAddNew:
            txt(1).Text = IndexAuto
            txt(1).Enabled = False
    End Select

End Sub

Private Sub DDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
                             ByVal pError As ADODB.Error, _
                             adStatus As ADODB.EventStatusEnum, _
                             ByVal pRecordset As ADODB.Recordset)
    Dim rsDetail As DBQuick
    Set rsDetail = New DBQuick
    rsDetail.DBOpen "select * from view_terima_produk where no_terima = '" & DDE.GetFieldByName("no_terima") & "'", CNN
    Set DDE.ChildRecordset = rsDetail.DBRecordset.Clone(adLockBatchOptimistic)
    Set DgDetail.DataSource = DDE.ChildRecordset
End Sub

Private Sub DDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb

        Case tmbSave:
            DDE.IsChildMemberReady = True
            simpan_header
            simpan_detail
    End Select

End Sub

Private Sub DGDETAIL_ButtonClick(ByVal ColIndex As Integer)
    rslot.DBOpen "select * from SHIEVER_HEADER", CNN
    Set lot = New frmCaller
    Set lot.FormData = rslot.DBRecordset
    lot.FromTagActive = "BLENDING INSTRUCTION"
    lot.CaptionLink = "BLENDING INSTRUCTION"
End Sub

Private Sub DGDETAIL_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        DDE.ChildRecordset.AddNew
    End If

End Sub

Private Sub Form_Load()

    With DDE
        Set .BindForm = Me
        .BindFormTAG = "terima"
        Set .ActiveConnection = CNN
        .PrepareQuery = "select * from T_PRODUK_HEADER"
    End With

    HiasForm Picture1, Me
    seting Me
End Sub

Function simpan_header()
    DDE.PrepareAppend = "insert into T_PRODUK_HEADER (no_terima,tanggal_terima) values ('" & txt(1).Text & "','" & Format(tgl.value, "yyyy-MM-dd") & "')"
    DDE.PrepareUpdate = "update T_PRODUK_HEADER  set tanggal_terima = '" & Format(tgl.value, "yyyy-MM-dd") & "' where no_terima = '" & txt(1).Text & "'"
End Function

Function simpan_detail()

    With DDE.ChildRecordset

        If .Recordcount <> 0 Then
            .MoveFirst

            If SendDataToServer(" delete from [T_PRODUK_DETAIL] where (no_terima = '" & DDE.GetFieldByName("no_terima") & "')") = True Then

                Do

                    If .EOF = True Then Exit Do
                    SendDataToServer "insert into T_PRODUK_DETAIL (no_terima,lot_no,jumlah,satuan)  " & " values ('" & txt(1).Text & "', " & " '" & .Fields("lot_no") & "', " & " '" & .Fields("jumlah") & "', " & " '" & .Fields("satuan") & "')"
                    .MoveNext
                Loop

            End If

            .MoveLast
            DgDetail.Refresh
        End If

    End With

End Function

Private Sub lot_RowColChange(ByVal TagForm As String, _
                             ByVal pRecordset As ADODB.Recordset)
    DDE.ChildRecordset.Fields("lot_no") = rslot.DBRecordset.Fields("lot_no")
End Sub
