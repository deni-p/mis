VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmprodpembungkusan 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pembungkusan"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9195
   Icon            =   "FrmProdPembungkusan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   9195
   Begin VB.TextBox lblEksno 
      Appearance      =   0  'Flat
      DataField       =   "noEkstraksi"
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
      Left            =   1320
      TabIndex        =   27
      Tag             =   "COV"
      Top             =   465
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4425
      Left            =   0
      ScaleHeight     =   4425
      ScaleWidth      =   9195
      TabIndex        =   8
      Top             =   0
      Width           =   9195
      Begin VB.TextBox txtwarnakain 
         Appearance      =   0  'Flat
         DataField       =   "warna_kain"
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
         Left            =   6840
         TabIndex        =   28
         Tag             =   "COV"
         Top             =   2400
         Width           =   2100
      End
      Begin VB.TextBox txtTotal 
         Appearance      =   0  'Flat
         DataField       =   "total"
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
         Left            =   6840
         TabIndex        =   23
         Tag             =   "COV"
         Top             =   3075
         Width           =   2100
      End
      Begin VB.TextBox txtHasilPembungkusan 
         Appearance      =   0  'Flat
         DataField       =   "hasil"
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
         Left            =   6840
         TabIndex        =   21
         Tag             =   "COV"
         Top             =   2730
         Width           =   2100
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         DataField       =   "Keterangan"
         DataSource      =   "MyDDE"
         Height          =   855
         Left            =   135
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Tag             =   "COV"
         Top             =   3450
         Width           =   8865
      End
      Begin VB.CommandButton cmdRefLink 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8610
         Picture         =   "FrmProdPembungkusan.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "BAHAN"
         Top             =   1335
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton cmdEkstraksi 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3405
         Picture         =   "FrmProdPembungkusan.frx":6BDC
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "Formula"
         Top             =   480
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtGroup 
         Appearance      =   0  'Flat
         DataField       =   "Group"
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
         Left            =   1320
         TabIndex        =   4
         Tag             =   "COV"
         Top             =   1155
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DcTanggal 
         DataField       =   "Tanggal"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Tag             =   "COV"
         Top             =   795
         Width           =   2055
         _ExtentX        =   3625
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
         Format          =   16515075
         CurrentDate     =   39634
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "waktu_mulai"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   0
         Left            =   6840
         TabIndex        =   6
         Tag             =   "COV"
         Top             =   1665
         Width           =   2100
         _ExtentX        =   3704
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
         CustomFormat    =   "dd MMM yyyy    HH:mm"
         Format          =   16515075
         CurrentDate     =   39419
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "waktu_selesai"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   1
         Left            =   6840
         TabIndex        =   7
         Tag             =   "COV"
         Top             =   2010
         Width           =   2100
         _ExtentX        =   3704
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
         CustomFormat    =   "dd MMM yyyy    HH:mm"
         Format          =   16515075
         CurrentDate     =   39419
      End
      Begin VB.Label lblGroup 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Warna Kain Pembungkusan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   4605
         TabIndex        =   29
         Top             =   2475
         Width           =   1950
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   10
         X1              =   7020
         X2              =   4575
         Y1              =   2700
         Y2              =   2700
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "approved_by"
         DataSource      =   "MyDDE"
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
         Left            =   6930
         TabIndex        =   26
         Top             =   120
         Width           =   2055
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   5
         X1              =   6970
         X2              =   5730
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label lblNoDokumen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   5730
         TabIndex        =   25
         Top             =   165
         Width           =   930
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   9
         X1              =   6990
         X2              =   4620
         Y1              =   3375
         Y2              =   3375
      End
      Begin VB.Label lblHasilPembungkusan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Stack"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   4605
         TabIndex        =   24
         Top             =   3150
         Width           =   795
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   7065
         X2              =   4620
         Y1              =   3030
         Y2              =   3030
      End
      Begin VB.Label lblGroup 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasil Pembungkusan (Lembar)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   4605
         TabIndex        =   22
         Top             =   2805
         Width           =   2145
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   7
         X1              =   7170
         X2              =   4620
         Y1              =   2310
         Y2              =   2310
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   6
         X1              =   7050
         X2              =   4620
         Y1              =   1965
         Y2              =   1965
      End
      Begin VB.Label lblTanggalWaktu 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal && waktu mulai"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4605
         TabIndex        =   19
         Top             =   1725
         Width           =   1890
      End
      Begin VB.Label lblReference 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4605
         TabIndex        =   17
         Top             =   1380
         Width           =   750
      End
      Begin VB.Label LbRefID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "refid"
         DataSource      =   "MyDDE"
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
         Height          =   330
         Left            =   6840
         TabIndex        =   16
         Tag             =   "COV"
         Top             =   1320
         Width           =   2100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   13
         X1              =   6870
         X2              =   4620
         Y1              =   1635
         Y2              =   1635
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   1360
         X2              =   120
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Label lblNoEkstraksi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Ekstraksi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   525
         Width           =   930
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   1360
         X2              =   120
         Y1              =   1095
         Y2              =   1095
      End
      Begin VB.Label lblid 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000005&
         Height          =   210
         Left            =   14640
         TabIndex        =   13
         Top             =   7560
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblTanggal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   870
         Width           =   570
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   1360
         X2              =   120
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lblGroup 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   1215
         Width           =   435
      End
      Begin VB.Label lblKeterangan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   165
         TabIndex        =   10
         Top             =   3195
         Width           =   945
      End
      Begin VB.Label lblNoDokumen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Dokumen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   165
         Width           =   960
      End
      Begin VB.Label lblDokNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "DokNo"
         DataSource      =   "MyDDE"
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
         Left            =   1320
         TabIndex        =   1
         Tag             =   "COV"
         Top             =   120
         Width           =   2055
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   8
         X1              =   1360
         X2              =   120
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label lblTanggalWaktu 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal && waktu selesai"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4605
         TabIndex        =   18
         Top             =   2100
         Width           =   2055
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   4440
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   1005
      BindFormTAG     =   "COV"
      ActiveLanguage  =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   1240
      X2              =   0
      Y1              =   255
      Y2              =   255
   End
   Begin VB.Label lblKeterangan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   585
   End
End
Attribute VB_Name = "frmprodpembungkusan"
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

Private Sub DetailProsesID(ByVal ParameterString As String)
    Set RcDetail = New DBQuick
    RcDetail.DBOpen "SELECT ProdFormulaEkstraksi.EksNo From ProdFormulaEkstraksi Where (ProdFormulaEkstraksi.typeTrans = '" & ParameterString & "')", CNN, lckLockBatch
End Sub

Private Sub OpenDetail(ByVal ParameterString As String)
    Dim ncount As Integer
    Set RcDetail = New DBQuick

    If ParameterString = "" Then ParameterString = "11111111111" ': Exit Sub
     
    RcDetail.DBOpen "SELECT  * from produksi_pembungkusan Where produksi_pembungkusan.DokNo = '" & ParameterString & "' ", CNN, lckLockBatch
    'Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
    'Set tgPembungkusan.DataSource = MyDDE.ChildRecordset

 
    RcDetail.CloseDB
    'isiGrid
End Sub

Private Sub PrepareQuery()
    On Error GoTo xErr
    Dim strSQL As String

    With MyDDE
        .PrepareAppend = "INSERT INTO produksi_pembungkusan(DokNo,NoEkstraksi,Tanggal,[Group],waktu_mulai,waktu_selesai,Keterangan,refid,hasil,total,type_proses,issued_by,warna_kain) VALUES('" & lblDokNo.Caption & "','" & lblEksno & "', CONVERT(DATETIME,'" & DcTanggal.Value & "',3),'" & txtGroup.Text & "', CONVERT(DATETIME,'" & tgl(0).Value & "',3),CONVERT(DATETIME,'" & tgl(1).Value & "',3),'" & txtKeterangan & "','" & LbRefID.Caption & "','" & txtHasilPembungkusan.Text & "','" & txtTotal.Text & "','PB','" & MainMenu.StatusBar1.Panels(1).Text & "','" & txtwarnakain.Text & "')"
        .PrepareUpdate = "UPDATE produksi_pembungkusan SET  NoEkstraksi = '" & lblEksno & "', Tanggal = CONVERT(DATETIME,'" & DcTanggal.Value & "',3), [Group] = '" & txtGroup.Text & "', waktu_mulai = CONVERT(DATETIME,'" & tgl(0).Value & "',3), waktu_selesai = CONVERT(DATETIME,'" & tgl(1).Value & "',3), Keterangan = '" & txtKeterangan.Text & "', refid = '" & LbRefID.Caption & "', type_proses = 'PB',hasil='" & txtHasilPembungkusan.Text & "',total='" & txtTotal.Text & "',warna_kain='" & txtwarnakain.Text & "'  where DokNo = '" & lblDokNo & "'"
        .PrepareDelete = " DELETE FROM  [produksi_pembungkusan] WHERE (DokNo = '" & .GetFieldByName("DokNo") & "')"
    End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error", msgOkOnly, msgExclamation
   Err.Clear
End Sub
Private Sub Form_Load()
   HiasFormManTell Picture2, Me
  
    With MyDDE
        .EditModeReplace = False
        Set .BindForm = Me
        .BindFormTAG = "COV"
        .SetPermissions = UserDeleteDenied
        Set .ActiveConnection = CNN
        .PrepareQuery = "select * from produksi_pembungkusan"
        .SetPermissions = aksess.MayDo("Pembungkusan")
    End With
End Sub

Private Sub lblEksno_LostFocus()
   Dim rsCek As New DBQuick
   rsCek.DBOpen "select * from statusProduksi where noEkstraksi='" & lblEksno.Text & "'", CNN, lckLockBatch
   If rsCek.DBRecordset.Recordcount > 0 Then
      rsCek.DBOpen "select * from produksi_pembungkusan where noEkstraksi='" & lblEksno.Text & "'", CNN, lckLockBatch
      If rsCek.DBRecordset.Recordcount > 0 Then
         MessageBox "Nomor Ekstraksi Ini Sudah Diinput...!", "Peringatan", msgOkOnly, msgCrtical
         lblEksno.Text = ""
      End If
   Else
      MessageBox "Nomor Ekstraksi Ini tidak ditemukan...!", "Peringatan", msgOkOnly, msgCrtical
      lblEksno.Text = ""
   End If
   rsCek.CloseDB
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
    Dim IDGen As New IDGenerator

    Select Case AdReasonActiveDb

        

        Case tmbAddNew

            MEdit = True
            Me.Tag = "baru"
            txtKeterangan.Text = "-"
            lblDokNo.Caption = IDGen.GetID("PB")
            cmdEkstraksi.Enabled = MEdit
            cmdRefLink.Enabled = True
            'lbleksno = frmProduksi.txtBox(1)
            LbRefID = frmProduksi.txtBox(5)
            txtGroup.SetFocus

            

        Case tmbCancel

            Me.Tag = ""

        Case tmbSave

            If MyDDE.IsChildMemberReady = True Then
                    SaveToMO
                    'updatePO
                    MEdit = False
                    If Me.Tag = "baru" Then
                        SendDataToServer "update statusproduksi set posisi='PEMBUNGKUSAN', status=1 where NoEkstraksi='" & lblEksno & "'"
                    End If
            End If
           OpenDetail lblDokNo.Caption
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
    Dim sWCID As String
   
    dStart = tgl(0).Value
    dFinish = tgl(1).Value
    ActualTime = Val(SelisihHariJam(dStart, dFinish, 2))
   
    rsCek.DBOpen "select WCID from WCenter_Header where FormID =43", CNN

    If rsCek.DBRecordset.Recordcount > 0 Then
        sWCID = rsCek.DBRecordset.Fields(0)
        SendDataToServer "update [order output detail] set actual_time=" & ActualTime & " where OrderID='" & LbRefID.Caption & "' and WCID='" & sWCID & "'"
    End If

    rsCek.CloseDB
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
                               ByVal pError As ADODB.Error, _
                               adStatus As ADODB.EventStatusEnum, _
                               ByVal pRecordset As ADODB.Recordset)
    If (MyDDE.ActiveRecordset.BOF = False) And (MyDDE.ActiveRecordset.EOF = False) Then OpenDetail IIf(IsNull(MyDDE.ActiveRecordset.Fields("DokNo")), "", MyDDE.ActiveRecordset.Fields("DokNo"))
    
   Label1.Caption = IIf(IsNull(MyDDE.GetFieldByName("approved_by")), "", MyDDE.GetFieldByName("approved_by"))
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb

        Case tmbAddNew

            MEdit = True
            'lbleksno = frmProduksi.txtBox(1)
            LbRefID = frmProduksi.txtBox(5)

        Case tmbSave
            If txtKeterangan.Text = "" Then txtKeterangan.Text = "-"
            If txtGroup.Text = "" Then txtGroup.Text = "-"
            If MyDDE.CheckEmptyControl = False Then
                    MyDDE.IsChildMemberReady = True
                    PrepareQuery
            Else
                MyDDE.IsChildMemberReady = False
            End If

        Case tmbDelete
            PrepareQuery

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

'Private Sub txtBarCode_KeyPress(KeyAscii As Integer)
'
'  If KeyAscii = 13 Then MaskSampling.SetFocus
'End Sub
'
'Private Sub txtBarCode_LostFocus()
'
'  If TxtBarcode.Text <> "" And Me.Tag = "AddNew" Then FindByTransID TxtBarcode.Text
'  Me.Tag = ""
'End Sub

Private Sub tgl_CallbackKeyDown(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub tgl_Change(Index As Integer)
   If tgl(0).Value > tgl(1).Value Then
      MessageBox "Waktu Mulai tidak boleh lebih besar dari waktu selesai", "Peringatan", msgOkOnly, msgCrtical
   End If
End Sub
