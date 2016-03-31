VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{B1E614FF-F86D-4F68-A86F-2584A0570C66}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmShiever 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   12000
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      ForeColor       =   &H80000008&
      Height          =   4620
      Left            =   -15
      ScaleHeight     =   4590
      ScaleWidth      =   12000
      TabIndex        =   1
      Top             =   -15
      Width           =   12030
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAAF6F&
         Caption         =   "SHIEVER POWDER LOT NO"
         Height          =   2220
         Left            =   135
         TabIndex        =   19
         Top             =   300
         Width           =   5820
         Begin VB.CommandButton cmdLink 
            Enabled         =   0   'False
            Height          =   330
            Left            =   3915
            MaskColor       =   &H000000C0&
            Picture         =   "frmShiever.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   25
            Tag             =   "SPPH"
            Top             =   300
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            DataField       =   "bag"
            DataSource      =   "DDE"
            Height          =   375
            Index           =   14
            Left            =   4845
            TabIndex        =   24
            Tag             =   "shiev"
            Top             =   1320
            Width           =   510
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            DataField       =   "kuantitas"
            DataSource      =   "DDE"
            Height          =   375
            Index           =   13
            Left            =   4845
            TabIndex        =   23
            Tag             =   "shiev"
            Top             =   1680
            Width           =   510
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            DataField       =   "residual_lot_no"
            DataSource      =   "DDE"
            Height          =   375
            Index           =   12
            Left            =   2160
            TabIndex        =   22
            Tag             =   "shiev"
            Top             =   1680
            Width           =   1005
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            DataField       =   "hasil"
            DataSource      =   "DDE"
            Height          =   375
            Index           =   11
            Left            =   2160
            TabIndex        =   21
            Tag             =   "shiev"
            Top             =   1320
            Width           =   1005
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            DataField       =   "lot_no"
            DataSource      =   "DDE"
            Height          =   375
            Index           =   16
            Left            =   2190
            TabIndex        =   20
            Tag             =   "shiev"
            Top             =   285
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "tgl_1"
            DataSource      =   "DDE"
            Height          =   345
            Index           =   2
            Left            =   2175
            TabIndex        =   26
            Tag             =   "shiev"
            Top             =   630
            Width           =   2580
            _ExtentX        =   4551
            _ExtentY        =   609
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd hh:mm:ss"
            Format          =   52166659
            CurrentDate     =   39423
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "tgl_2"
            DataSource      =   "DDE"
            Height          =   345
            Index           =   3
            Left            =   2160
            TabIndex        =   27
            Tag             =   "shiev"
            Top             =   975
            Width           =   2580
            _ExtentX        =   4551
            _ExtentY        =   609
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd hh:mm:ss"
            Format          =   52166659
            CurrentDate     =   39423
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Kg"
            Height          =   255
            Index           =   25
            Left            =   5415
            TabIndex        =   36
            Top             =   1770
            Width           =   285
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Kg"
            Height          =   255
            Index           =   24
            Left            =   5430
            TabIndex        =   35
            Top             =   1410
            Width           =   285
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Bag"
            Height          =   255
            Index           =   23
            Left            =   3255
            TabIndex        =   34
            Top             =   1410
            Width           =   360
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Kuantitas sisa powder"
            Height          =   255
            Index           =   19
            Left            =   3225
            TabIndex        =   33
            Top             =   1755
            Width           =   1650
         End
         Begin VB.Line Line2 
            Index           =   4
            X1              =   300
            X2              =   5235
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Line Line2 
            Index           =   3
            X1              =   285
            X2              =   5220
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Line Line2 
            Index           =   2
            X1              =   285
            X2              =   2625
            Y1              =   1305
            Y2              =   1305
         End
         Begin VB.Line Line2 
            Index           =   1
            X1              =   255
            X2              =   2595
            Y1              =   645
            Y2              =   645
         End
         Begin VB.Line Line2 
            Index           =   0
            X1              =   255
            X2              =   2595
            Y1              =   975
            Y2              =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Residual Lot No"
            Height          =   255
            Index           =   17
            Left            =   300
            TabIndex        =   32
            Top             =   1740
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Hasil Powder"
            Height          =   255
            Index           =   16
            Left            =   285
            TabIndex        =   31
            Top             =   1395
            Width           =   1005
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal dan waktu mulai"
            Height          =   255
            Index           =   15
            Left            =   255
            TabIndex        =   30
            Top             =   1035
            Width           =   2490
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal dan waktu mulai"
            Height          =   255
            Index           =   14
            Left            =   255
            TabIndex        =   29
            Top             =   705
            Width           =   2490
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Lot No"
            Height          =   255
            Index           =   13
            Left            =   255
            TabIndex        =   28
            Top             =   390
            Width           =   645
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Operator Shiever"
         Height          =   2235
         Index           =   0
         Left            =   6120
         TabIndex        =   17
         Top             =   285
         Width           =   5670
         Begin MSDataGridLib.DataGrid dgdetail 
            Bindings        =   "frmShiever.frx":038A
            Height          =   1860
            Left            =   225
            TabIndex        =   18
            Tag             =   "shiev"
            Top             =   300
            Width           =   5235
            _ExtentX        =   9234
            _ExtentY        =   3281
            _Version        =   393216
            AllowUpdate     =   -1  'True
            Enabled         =   0   'False
            HeadLines       =   1
            RowHeight       =   15
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
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "nama_shiever"
               Caption         =   "Operator Shiever"
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
               Caption         =   "Jumlah"
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
               Caption         =   "Satuan"
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
                  Button          =   -1  'True
               EndProperty
               BeginProperty Column01 
               EndProperty
               BeginProperty Column02 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Checklist Kebersihan Blend Powder"
         Height          =   1770
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   2640
         Width           =   5865
         Begin VB.Frame Frame1 
            BackColor       =   &H00EAAF6F&
            BorderStyle     =   0  'None
            Height          =   345
            Index           =   0
            Left            =   4170
            TabIndex        =   10
            Top             =   450
            Width           =   1575
            Begin VB.OptionButton Option1 
               Appearance      =   0  'Flat
               BackColor       =   &H00EAAF6F&
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   0
               Left            =   300
               TabIndex        =   12
               Top             =   90
               Width           =   240
            End
            Begin VB.OptionButton Option1 
               Appearance      =   0  'Flat
               BackColor       =   &H00EAAF6F&
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   1
               Left            =   1020
               TabIndex        =   11
               Top             =   90
               Width           =   240
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00EAAF6F&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   1
            Left            =   4155
            TabIndex        =   7
            Top             =   810
            Width           =   1575
            Begin VB.OptionButton Option2 
               Appearance      =   0  'Flat
               BackColor       =   &H00EAAF6F&
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   1
               Left            =   1020
               TabIndex        =   9
               Top             =   15
               Width           =   240
            End
            Begin VB.OptionButton Option2 
               Appearance      =   0  'Flat
               BackColor       =   &H00EAAF6F&
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   0
               Left            =   300
               TabIndex        =   8
               Top             =   15
               Width           =   240
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00EAAF6F&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   2
            Left            =   4155
            TabIndex        =   4
            Top             =   1140
            Width           =   1575
            Begin VB.OptionButton Option3 
               Appearance      =   0  'Flat
               BackColor       =   &H00EAAF6F&
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   1
               Left            =   1005
               TabIndex        =   6
               Top             =   0
               Width           =   240
            End
            Begin VB.OptionButton Option3 
               Appearance      =   0  'Flat
               BackColor       =   &H00EAAF6F&
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   0
               Left            =   285
               TabIndex        =   5
               Top             =   0
               Width           =   240
            End
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama                                                                                    Bersih      Kotor"
            Height          =   255
            Index           =   12
            Left            =   210
            TabIndex        =   16
            Top             =   255
            Width           =   5415
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Bagian dalam mesin blender"
            Height          =   255
            Index           =   26
            Left            =   195
            TabIndex        =   15
            Top             =   885
            Width           =   2490
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Lantai"
            Height          =   255
            Index           =   27
            Left            =   195
            TabIndex        =   14
            Top             =   540
            Width           =   645
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   5445
            X2              =   165
            Y1              =   1125
            Y2              =   1125
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Bagian luar mesin blender"
            Height          =   255
            Index           =   28
            Left            =   195
            TabIndex        =   13
            Top             =   1185
            Width           =   2490
         End
         Begin VB.Line Line1 
            Index           =   13
            X1              =   5460
            X2              =   165
            Y1              =   1425
            Y2              =   1425
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   5460
            X2              =   180
            Y1              =   795
            Y2              =   795
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "id_blending"
         DataSource      =   "DDE"
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Tag             =   "shiev"
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   4590
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   1005
      BindFormTAG     =   "mixing"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmShiever"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents lotno As frmCaller
Attribute lotno.VB_VarHelpID = -1
Private WithEvents op_shiever As frmCaller
Attribute op_shiever.VB_VarHelpID = -1
Dim rsop As New DBQuick
Dim rslotno As New DBQuick
Dim lantai As String
Dim dalam As String
Dim luar As String
Private Sub cmdLink_Click()
   rslotno.DBOpen "select * from BLENDING_INSTRUCTION", CNN
   Set lotno = New frmCaller
   Set lotno.FormData = rslotno.DBRecordset
   lotno.FromTagActive = "BLENDING INSTRUCTION"
   lotno.CaptionLink = "BLENDING INSTRUCTION"
End Sub


Private Sub DDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
Case tmbAddNew:
    cmdLink.Enabled = True
    DGDETAIL.Enabled = True
    Option1(0).value = False
    Option1(1).value = False
    Option2(0).value = False
    Option2(1).value = False
    Option3(0).value = False
    Option3(1).value = False
End Select
End Sub

Private Sub DDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim rsDetail As DBQuick
Set rsDetail = New DBQuick
If DDE.GetFieldByName("lantai") = "Bersih" Then Option1(0).value = True
If DDE.GetFieldByName("lantai") = "Kotor" Then Option1(1).value = True
If DDE.GetFieldByName("luar") = "Bersih" Then Option2(0).value = True
If DDE.GetFieldByName("luar") = "Kotor" Then Option2(1).value = True
If DDE.GetFieldByName("dalam") = "Bersih" Then Option3(0).value = True
If DDE.GetFieldByName("dalam") = "Kotor" Then Option3(1).value = True

rsDetail.DBOpen "select * from VIEW_SHIEVER where lot_no = '" & DDE.GetFieldByName("lot_no") & "'", CNN
Set DDE.ChildRecordset = rsDetail.DBRecordset.Clone(adLockBatchOptimistic)
Set DGDETAIL.DataSource = DDE.ChildRecordset


End Sub


Private Sub DDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
Case tmbSave:
     DDE.IsChildMemberReady = True
     simpan
     simpan_detail
End Select
End Sub

Function simpan()
DDE.PrepareAppend = "insert into SHIEVER_HEADER (id_blending,lot_no,tgl_1,tgl_2,hasil,bag,residual_lot_no,kuantitas,lantai,dalam,luar) values ('" & Text1(0).Text & "','" & Text1(16).Text & "' , '" & Format(DTPicker1(2).value, "yyyy-MM-dd hh:mm:ss") & "', '" & Format(DTPicker1(3).value, "yyyy-MM-dd hh:mm:ss") & "', '" & Text1(11).Text & "', '" & Text1(14).Text & "', '" & Text1(12).Text & "', '" & Text1(13).Text & "', '" & lantai & "','" & dalam & "', '" & luar & "')"

DDE.PrepareUpdate = "update SHIEVER_HEADER set tgl_1 = '" & Format(DTPicker1(2).value, "yyyy-MM-dd hh:mm:ss") & "', tgl_2 = '" & Format(DTPicker1(3).value, "yyyy-MM-dd hh:mm:ss") & "', hasil = '" & Text1(11).Text & "', bag = '" & Text1(14).Text & "', residual_lot_no = '" & Text1(12).Text & "', kuantitas = '" & Text1(13).Text & "', lantai = '" & lantai & "', dalam = '" & dalam & "', luar = '" & luar & "' where id_blending = '" & Text1(0).Text & "'"

End Function

Private Sub DGDETAIL_ButtonClick(ByVal ColIndex As Integer)
Select Case ColIndex
Case 0
    rsop.DBOpen "select * from OPERATOR_SHIEVER", CNN
    Set op_shiever = New frmCaller
    Set op_shiever.FormData = rsop.DBRecordset
    op_shiever.FromTagActive = "OPERATOR SHIEVER"
    op_shiever.CaptionLink = "OPERATOR SHIEVER"
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
    .BindFormTAG = "shiev"
Set .ActiveConnection = CNN
    .PrepareQuery = "select * from SHIEVER_HEADER"
End With

HiasForm Picture2, Me
seting Me
End Sub
Private Sub lotno_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Text1(0).Text = rslotno.DBRecordset.Fields("id_blending")
Text1(16).Text = rslotno.DBRecordset.Fields("lot_no")
End Sub

Private Sub op_shiever_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
DDE.ChildRecordset.Fields("nama_shiever") = rsop.DBRecordset.Fields("nama_shiever")
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
      lantai = "Bersih"
Case 1
      lantai = "Kotor"
End Select
End Sub

Function simpan_detail()
With DDE.ChildRecordset
   If .Recordcount <> 0 Then
       .MoveFirst
         If SendDataToServer(" delete from [SHIEVER_DETAIL] where (lot_no = '" & DDE.GetFieldByName("lot_no") & "')") = True Then
         Do
           If .EOF = True Then Exit Do
           SendDataToServer "insert into SHIEVER_DETAIL (lot_no,nama_shiever,jumlah,satuan)  " & _
           " values ('" & Text1(16).Text & "', " & _
           " '" & .Fields("nama_shiever") & "', " & _
           " '" & .Fields("jumlah") & "', " & _
           " '" & .Fields("satuan") & "')"
          .MoveNext
        Loop
        End If
        .MoveLast
        DGDETAIL.Refresh
        End If
    End With
End Function

Private Sub Option2_Click(Index As Integer)
Select Case Index
Case 0
      dalam = "Bersih"
Case 1
      dalam = "Kotor"
End Select
End Sub

Private Sub Option3_Click(Index As Integer)
Select Case Index
Case 0
      luar = "Bersih"
Case 1
      luar = "Kotor"
End Select
End Sub

