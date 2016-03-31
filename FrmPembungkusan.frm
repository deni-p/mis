VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{73779082-7BF1-482D-A01F-0D9823B548F1}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmPembungkusan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pembungkusan"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10440
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   10440
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      ForeColor       =   &H80000008&
      Height          =   2865
      Left            =   15
      ScaleHeight     =   2835
      ScaleWidth      =   10395
      TabIndex        =   7
      Top             =   0
      Width           =   10425
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "total"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   5
         Left            =   7170
         TabIndex        =   20
         Tag             =   "bungkus"
         Top             =   1785
         Width           =   1695
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "kdwarna"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   4
         Left            =   7170
         TabIndex        =   18
         Tag             =   "bungkus"
         Top             =   1185
         Width           =   1695
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4020
         MaskColor       =   &H000000C0&
         Picture         =   "FrmPembungkusan.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "SPPH"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "hasil_bungkus"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   3
         Left            =   7170
         TabIndex        =   5
         Tag             =   "bungkus"
         Top             =   1485
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker tgl_mulai 
         DataField       =   "tanggal_mulai"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "m/d/yy h:nn AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   7155
         TabIndex        =   3
         Tag             =   "bungkus"
         Top             =   555
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy hh:mm:ss"
         Format          =   60555267
         CurrentDate     =   39401
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "ket_bungkus"
         DataSource      =   "DDE"
         Height          =   555
         Index           =   40
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Tag             =   "bungkus"
         Top             =   2040
         Width           =   3270
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "grup"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   2
         Tag             =   "bungkus"
         Top             =   885
         Width           =   1350
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "No_ekstraksi"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   2280
         TabIndex        =   0
         Tag             =   "bungkus"
         Top             =   120
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker tgl_bungkus 
         DataField       =   "tanggal_bungkus"
         DataSource      =   "DDE"
         Height          =   315
         Left            =   2280
         TabIndex        =   1
         Tag             =   "bungkus"
         Top             =   570
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   60555267
         CurrentDate     =   39365
      End
      Begin MSComCtl2.DTPicker tgl_selesai 
         DataField       =   "tanggal_selesai"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "m/d/yy h:nn AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   7155
         TabIndex        =   4
         Tag             =   "bungkus"
         Top             =   870
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy hh:mm:ss"
         Format          =   60555267
         CurrentDate     =   39401
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Stack (stack)"
         Height          =   255
         Index           =   7
         Left            =   4995
         TabIndex        =   6
         Top             =   1860
         Width           =   1665
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   7560
         X2              =   4980
         Y1              =   2085
         Y2              =   2085
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Warna Kain"
         Height          =   255
         Index           =   6
         Left            =   4980
         TabIndex        =   19
         Top             =   1260
         Width           =   1665
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   7545
         X2              =   4965
         Y1              =   1485
         Y2              =   1485
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   7560
         X2              =   4980
         Y1              =   1785
         Y2              =   1785
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   7560
         X2              =   4980
         Y1              =   1185
         Y2              =   1185
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   7545
         X2              =   4965
         Y1              =   870
         Y2              =   870
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Hasil Pembungkusan"
         Height          =   255
         Index           =   5
         Left            =   4995
         TabIndex        =   16
         Top             =   1560
         Width           =   1665
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Dan Waktu Selesai"
         Height          =   255
         Index           =   4
         Left            =   4980
         TabIndex        =   15
         Top             =   975
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Dan Waktu Mulai"
         Height          =   255
         Index           =   3
         Left            =   4980
         TabIndex        =   14
         Top             =   645
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   480
         TabIndex        =   13
         Top             =   1725
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "No Ekstrasi"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   180
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   645
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   2055
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2400
         X2              =   360
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2400
         X2              =   360
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   2400
         X2              =   360
         Y1              =   1185
         Y2              =   1185
      End
   End
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   11
      Top             =   2865
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   1005
      BindFormTAG     =   "TTRL"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmPembungkusan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private WithEvents Bleacing  As frmCaller
Attribute Bleacing.VB_VarHelpID = -1
Dim rsbleacing As New DBQuick

Private Sub Bleacing_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
   If rsbleacing.DBRecordset.Recordcount > 0 Then
      txt(0).Text = rsbleacing.DBRecordset.Fields("no_ekstrasi")
   End If
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
    cmdLink.Enabled = True
End Select
End Sub

Private Sub DDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
 Case tmbSave:
      DDE.IsChildMemberReady = True
      simpan
 Case tmbDelete:
      DDE.IsChildMemberReady = True
      Del
End Select
End Sub
Function simpan()
DDE.PrepareAppend = "insert into PEMBUNGKUSAN (no_ekstraksi,tanggal_bungkus,grup,tanggal_mulai,tanggal_selesai,hasil_bungkus,ket_bungkus) values ('" + DDE.GetFieldByName("no_ekstraksi") + "', '" + Format(tgl_bungkus.value, "yyyy-MM-dd hh:mm:ss") + "','" + DDE.GetFieldByName("grup") + "', '" + Format(tgl_mulai(0).value, "yyyy-MM-dd hh:mm:ss") + "','" + Format(tgl_selesai(1).value, "yyyy-MM-dd hh:mm:ss") + "','" + DDE.GetFieldByName("hasil_bungkus") + "','" + DDE.GetFieldByName("ket_bungkus") + "')"
DDE.PrepareUpdate = "update PEMBUNGKUSAN set tanggal_bungkus = '" & Format(tgl_bungkus.value, "yyyy-MM-dd hh:mm:ss") & ", Grup = '" & DDE.GetFieldByName("Grup") & "', tgl_mulai = '" & Format(tgl_mulai(0).value, "yyyy-MM-dd hh:mm:ss") & "', " & _
                    " tgl_selesai = '" & Format(tgl_selesai(1).value, "yyyy-MM-dd hh:mm:ss") & "', " & _
                    " hasil_bungkus = '" & DDE.GetFieldByName("hasil_bungkus") & "', " & _
                    " keterangan_bungkus = '" & DDE.GetFieldByName("keterangan_bungkus") & "' where no_ekstrasi ='" & DDE.GetFieldByName("no_ekstrasi") & "'"
End Function
Function Del()
DDE.PrepareDelete = "delete from PEMBUNGKUSAN where no_ekstraksi = '" + DDE.GetFieldByName("no_ekstraksi") + "'"
End Function

Private Sub Form_Load()
   With DDE
   Set .BindForm = Me
       .BindFormTAG = "bungkus"
   Set .ActiveConnection = CNN
       .PrepareQuery = "select * from PEMBUNGKUSAN"
   End With
   HiasFormManTell Picture2, Me
   seting Me
End Sub

Private Sub txt_GotFocus(Index As Integer)
txt(Index).BackColor = &H79BCFF
End Sub

Private Sub txt_LostFocus(Index As Integer)
txt(Index).BackColor = vbWhite
End Sub
