VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{D7BB8F75-AC9E-4E80-A526-70EA20ACFD16}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmfilter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter Press"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11745
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   11745
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5985
      Left            =   0
      ScaleHeight     =   5985
      ScaleWidth      =   11745
      TabIndex        =   1
      Top             =   0
      Width           =   11745
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "no_ekstrasi"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   1
         Left            =   1200
         TabIndex        =   7
         Tag             =   "filter"
         Top             =   420
         Width           =   1695
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "id_filter"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   6
         Tag             =   "filter"
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "grup"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   2
         Left            =   1185
         TabIndex        =   5
         Tag             =   "filter"
         Top             =   1020
         Width           =   1710
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "desk_filter"
         DataSource      =   "DDE"
         Height          =   915
         Index           =   3
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Tag             =   "filter"
         Top             =   4860
         Width           =   4245
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1005
         ItemData        =   "fpress.frx":0000
         Left            =   375
         List            =   "fpress.frx":0002
         TabIndex        =   3
         Top             =   2355
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2910
         MaskColor       =   &H000000C0&
         Picture         =   "fpress.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "SPPH"
         Top             =   435
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin MSDataGridLib.DataGrid DGDETAIL 
         Height          =   3000
         Left            =   75
         TabIndex        =   8
         Tag             =   "filter"
         Top             =   1500
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   5292
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         Appearance      =   0
         DefColWidth     =   6667
         HeadLines       =   2
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
         Caption         =   "FILTER PRESS"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "nama_proses"
            Caption         =   "PROSES"
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
            DataField       =   "a"
            Caption         =   "MULAI POMPA"
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
            DataField       =   "b"
            Caption         =   "SELESAI POMPA"
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
            DataField       =   "c"
            Caption         =   "MULAI GANTI KAIN"
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
            DataField       =   "d"
            Caption         =   "SELESAI GANTI KAIN"
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
            DataField       =   "e"
            Caption         =   "MULAI BONGKAR FILTER PRESS"
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
         BeginProperty Column06 
            DataField       =   "f"
            Caption         =   "SELESAI BONGKAR FILTER PRESS"
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
               Button          =   -1  'True
            EndProperty
            BeginProperty Column02 
               Button          =   -1  'True
               WrapText        =   -1  'True
            EndProperty
            BeginProperty Column03 
               Button          =   -1  'True
            EndProperty
            BeginProperty Column04 
               Button          =   -1  'True
            EndProperty
            BeginProperty Column05 
               Button          =   -1  'True
            EndProperty
            BeginProperty Column06 
               Button          =   -1  'True
               ColumnWidth     =   1995,024
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "tanggal_press"
         DataSource      =   "DDE"
         Height          =   315
         Left            =   1185
         TabIndex        =   9
         Tag             =   "filter"
         Top             =   705
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   62062595
         CurrentDate     =   39365
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   2880
         X2              =   105
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2130
         X2              =   90
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   255
         Index           =   2
         Left            =   105
         TabIndex        =   14
         Top             =   780
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "No Ekstrasi"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   495
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Filter Press"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   195
         Width           =   2055
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2145
         X2              =   105
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   11
         Top             =   1095
         Width           =   2055
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   2145
         X2              =   105
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         Height          =   255
         Index           =   4
         Left            =   90
         TabIndex        =   10
         Top             =   4560
         Width           =   2055
      End
   End
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   6000
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   1005
      BindFormTAG     =   "cruz"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmfilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents Bleacing  As frmCaller
Attribute Bleacing.VB_VarHelpID = -1
Dim rsbleacing As New DBQuick
Dim data As Integer

Private Sub Bleacing_RowColChange(ByVal TagForm As String, _
                                  ByVal pRecordset As ADODB.Recordset)
    txt(1).Text = rsbleacing.DBRecordset.Fields("no_ekstrasi")
End Sub

Private Sub cmdLink_Click()
    rsbleacing.DBOpen "select * from BLEACHING", CNN

    If rsbleacing.DBRecordset.EOF Then
        rsbleacing.DBOpen "select * from ACID_TREATMEN", CNN
    Else
        rsbleacing.DBOpen "select * from BLEACHING, ACID_TREATMEN where BLEACHING.no_ekstrasi <> ACID_TREATMEN.no_ekstrasi ", CNN
    End If

    Set Bleacing = New frmCaller
    Set Bleacing.FormData = rsbleacing.DBRecordset
    Bleacing.FromTagActive = "BLEACHING TREATMEN"
    Bleacing.CaptionLink = "BLEACHING TREATMEN"
End Sub

Private Sub DDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb

        Case tmbAddNew:
            CmdLink.Enabled = True
            txt(0).Text = IndexAuto
            txt(0).Enabled = False
    
        Case tmbSave
            DDE.IsChildMemberReady = True
            simpan_header
            simpan_detail
    
    End Select

End Sub

Private Sub DDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
                             ByVal pError As ADODB.Error, _
                             adStatus As ADODB.EventStatusEnum, _
                             ByVal pRecordset As ADODB.Recordset)
    Dim rsDetail As New DBQuick
    Set rsDetail = New DBQuick
    rsDetail.DBOpen "select * from view_filter where id_filter = '" & DDE.GetFieldByName("id_filter") & "'", CNN
    Set DDE.ChildRecordset = rsDetail.DBRecordset.Clone(adLockBatchOptimistic)
    Set DgDetail.DataSource = DDE.ChildRecordset
End Sub

Private Sub DDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb
    
        Case tmbDelete
            delete
    End Select

End Sub

Private Sub DGDETAIL_ButtonClick(ByVal ColIndex As Integer)

    Select Case ColIndex

        Case 0
            List1.Clear
            List1.AddItem "Filter Press I"
            List1.AddItem "Filter Press II"
            List1.AddItem "Filter Press III"
            List1.AddItem "Filter Press IV"
            List1.AddItem "Filter Press V"
            List1.Visible = True
            List1.Move DgDetail.Columns(0).Left + 100, (DgDetail.RowTop(DgDetail.Row) + DgDetail.Top)
    End Select

End Sub

Private Sub DGDETAIL_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        DDE.ChildRecordset.AddNew
    End If

End Sub

Private Sub Form_Load()

    With DDE
        Set .BindForm = Me
        .BindFormTAG = "filter"
        Set .ActiveConnection = CNN
        .PrepareQuery = "select * from filter_header"
    End With

    'HiasForm Picture1, Me
End Sub

Function delete()
    DDE.PrepareDelete = "delete from filter_header where id_filter = '" & txt(0).Text & "'"
End Function

Function simpan_header()
    DDE.PrepareAppend = "insert into filter_header (id_filter, no_ekstrasi, tanggal_press,grup,desk_filter) values ('" & txt(0).Text & "', '" & DDE.GetFieldByName("no_ekstrasi") & "','" & Format(tgl.value, "yyyy-MM-dd") & "', '" & DDE.GetFieldByName("grup") & "', '" & DDE.GetFieldByName("desk_filter") & "')"
    DDE.PrepareUpdate = " update filter_header set id_filter = '" & txt(0).Text & "', no_ekstrasi = '" & DDE.GetFieldByName("no_ekstrasi") & "', tanggal_press = '" & Format(tgl.value, "yyyy-MM-dd") & "', grup = '" & DDE.GetFieldByName("grup") & "', desk_filter = '" & DDE.GetFieldByName("desk_filter") & "' where id_filter = '" & txt(0).Text & "'"
End Function

Function simpan_detail()

    With DDE.ChildRecordset

        If .Recordcount <> 0 Then
            .MoveFirst

            If SendDataToServer(" delete from [filter_detail] where (id_filter = '" & DDE.GetFieldByName("id_filter") & "')") = True Then

                Do

                    If .EOF = True Then Exit Do
                    SendDataToServer "insert into filter_detail (id_filter,nama_proses,a,b,c,d,e,f)  " & " values ('" & txt(0).Text & "', " & " '" & .Fields("nama_proses") & "', " & " '" & .Fields("a") & "', " & " '" & .Fields("b") & "', " & " '" & .Fields("c") & "', " & " '" & .Fields("d") & "', " & " '" & .Fields("e") & "', " & " '" & .Fields("f") & "')"
                    .MoveNext
                Loop

            End If

            .MoveLast
            DgDetail.Refresh
        End If

    End With

End Function

Private Function IndexAuto() As String
    Dim Rc As New DBQuick
    Dim TglSaiki As String
    Dim Inom As Long
    TglSaiki = Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2)
    Rc.DBOpen "SELECT MAX(RIGHT(id_filter, 5)) AS MaxNom FROM [filter_header] WHERE (GETDATE() = { fn NOW() })", CNN, lckLockReadOnly

    With Rc

        If .DBRecordset.Recordcount <> 0 Then
            Inom = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + 1
        Else
            Inom = 1
        End If

        Select Case Len(Trim(Str(Inom)))

            Case 0: IndexAuto = "FP/" & TglSaiki & "-" & Trim(Str(Inom))

            Case 1: IndexAuto = "FP/" & TglSaiki & "-" & "0000" & Trim(Str(Inom))

            Case 2: IndexAuto = "FP/" & TglSaiki & "-" & "000" & Trim(Str(Inom))

            Case 3: IndexAuto = "FP/" & TglSaiki & "-" & "00" & Trim(Str(Inom))

            Case 4: IndexAuto = "FP/" & TglSaiki & "-" & "0" & Trim(Str(Inom))
        End Select

    End With

End Function

Private Sub List1_Click()
    DDE.ChildRecordset.Fields("nama_proses") = List1.Text
    List1.Visible = False
End Sub
