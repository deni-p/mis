VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmLembarSupplier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lembar Supplier"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLembarSupplier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   9135
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6885
      Left            =   0
      ScaleHeight     =   6885
      ScaleWidth      =   9135
      TabIndex        =   1
      Top             =   0
      Width           =   9135
      Begin VB.TextBox txtSupplier 
         Appearance      =   0  'Flat
         DataField       =   "companyName"
         Height          =   315
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   67
         Tag             =   "LS"
         Top             =   225
         Width           =   2145
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "qty_receive"
         Height          =   315
         Left            =   1650
         TabIndex        =   66
         Tag             =   "LS"
         Top             =   930
         Width           =   1380
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   3795
         Picture         =   "frmLembarSupplier.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   64
         Tag             =   "jemur"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtKet 
         Appearance      =   0  'Flat
         DataField       =   "keterangan"
         Height          =   690
         Left            =   270
         MultiLine       =   -1  'True
         TabIndex        =   62
         Tag             =   "LS"
         Top             =   6135
         Width           =   8610
      End
      Begin VB.TextBox txtBeratKirim 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "berat_kirim3"
         Height          =   315
         Index           =   2
         Left            =   4905
         TabIndex        =   58
         Tag             =   "LS"
         Top             =   5385
         Width           =   1170
      End
      Begin VB.TextBox txtBeratKirim 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "berat_kirim2"
         Height          =   315
         Index           =   1
         Left            =   4905
         TabIndex        =   57
         Tag             =   "LS"
         Top             =   5055
         Width           =   1170
      End
      Begin VB.TextBox txtBeratKirim 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "berat_kirim1"
         Height          =   315
         Index           =   0
         Left            =   4905
         TabIndex        =   56
         Tag             =   "LS"
         Top             =   4725
         Width           =   1170
      End
      Begin VB.TextBox txtBeratProses 
         Appearance      =   0  'Flat
         DataField       =   "berat_proses"
         Height          =   315
         Left            =   1875
         TabIndex        =   48
         Tag             =   "LS"
         Top             =   4245
         Width           =   1050
      End
      Begin VB.TextBox txtHari 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "hari_packing"
         Height          =   315
         Index           =   4
         Left            =   6345
         TabIndex        =   39
         Tag             =   "LS"
         Top             =   3195
         Width           =   1380
      End
      Begin VB.TextBox txtHari 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "hari_napel"
         Height          =   315
         Index           =   3
         Left            =   6345
         TabIndex        =   38
         Tag             =   "LS"
         Top             =   2850
         Width           =   1380
      End
      Begin VB.TextBox txtHari 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "hari_sortir"
         Height          =   315
         Index           =   2
         Left            =   6345
         TabIndex        =   37
         Tag             =   "LS"
         Top             =   2505
         Width           =   1380
      End
      Begin VB.TextBox txtHari 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "hari_jemur"
         Height          =   315
         Index           =   1
         Left            =   6345
         TabIndex        =   36
         Tag             =   "LS"
         Top             =   2160
         Width           =   1380
      End
      Begin VB.TextBox txtHari 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "hari_cuci"
         Height          =   315
         Index           =   0
         Left            =   6345
         TabIndex        =   35
         Tag             =   "LS"
         Top             =   1815
         Width           =   1380
      End
      Begin MSComCtl2.DTPicker DTPProses 
         DataField       =   "tgl_cuci"
         Height          =   315
         Index           =   0
         Left            =   4260
         TabIndex        =   30
         Tag             =   "LS"
         Top             =   1800
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   61603843
         CurrentDate     =   39651
      End
      Begin VB.TextBox txtJmlOrang 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "org_packing"
         Height          =   315
         Index           =   4
         Left            =   1425
         TabIndex        =   18
         Tag             =   "LS"
         Top             =   3180
         Width           =   1230
      End
      Begin VB.TextBox txtJmlOrang 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "org_napel"
         Height          =   315
         Index           =   3
         Left            =   1425
         TabIndex        =   17
         Tag             =   "LS"
         Top             =   2835
         Width           =   1230
      End
      Begin VB.TextBox txtJmlOrang 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "org_sortir"
         Height          =   315
         Index           =   2
         Left            =   1425
         TabIndex        =   16
         Tag             =   "LS"
         Top             =   2490
         Width           =   1230
      End
      Begin VB.TextBox txtJmlOrang 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "org_jemur"
         Height          =   315
         Index           =   1
         Left            =   1425
         TabIndex        =   15
         Tag             =   "LS"
         Top             =   2145
         Width           =   1230
      End
      Begin VB.TextBox txtJmlOrang 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "org_cuci"
         Height          =   315
         Index           =   0
         Left            =   1425
         TabIndex        =   14
         Tag             =   "LS"
         Top             =   1800
         Width           =   1230
      End
      Begin VB.CheckBox chkProses 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Packing"
         DataField       =   "packing"
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   4
         Left            =   450
         TabIndex        =   13
         Tag             =   "LS"
         Top             =   3165
         Width           =   1020
      End
      Begin VB.CheckBox chkProses 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Napel"
         DataField       =   "napel"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   3
         Left            =   450
         TabIndex        =   12
         Tag             =   "LS"
         Top             =   2820
         Width           =   780
      End
      Begin VB.CheckBox chkProses 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Sortir"
         DataField       =   "sortir"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   450
         TabIndex        =   11
         Tag             =   "LS"
         Top             =   2475
         Width           =   780
      End
      Begin VB.CheckBox chkProses 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Jemur"
         DataField       =   "jemur"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   450
         TabIndex        =   10
         Tag             =   "LS"
         Top             =   2145
         Width           =   780
      End
      Begin VB.CheckBox chkProses 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Cuci"
         DataField       =   "cuci"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   450
         TabIndex        =   9
         Tag             =   "LS"
         Top             =   1800
         Width           =   780
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Kondisi Rumput Laut"
         Height          =   735
         Left            =   6450
         TabIndex        =   6
         Top             =   480
         Width           =   2475
         Begin VB.OptionButton Op 
            BackColor       =   &H00EAAF6F&
            Caption         =   "Kering"
            Height          =   285
            Index           =   1
            Left            =   1260
            TabIndex        =   8
            Tag             =   "LS"
            Top             =   315
            Width           =   960
         End
         Begin VB.OptionButton Op 
            BackColor       =   &H00EAAF6F&
            Caption         =   "Basah"
            Height          =   285
            Index           =   0
            Left            =   105
            TabIndex        =   7
            Tag             =   "LS"
            Top             =   315
            Width           =   1260
         End
      End
      Begin MSComCtl2.DTPicker DTPProses 
         DataField       =   "tgl_jemur"
         Height          =   315
         Index           =   1
         Left            =   4260
         TabIndex        =   31
         Tag             =   "LS"
         Top             =   2145
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   61603843
         CurrentDate     =   39651
      End
      Begin MSComCtl2.DTPicker DTPProses 
         DataField       =   "tgl_sortir"
         Height          =   315
         Index           =   2
         Left            =   4260
         TabIndex        =   32
         Tag             =   "LS"
         Top             =   2490
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   61603843
         CurrentDate     =   39651
      End
      Begin MSComCtl2.DTPicker DTPProses 
         DataField       =   "tgl_napel"
         Height          =   315
         Index           =   3
         Left            =   4260
         TabIndex        =   33
         Tag             =   "LS"
         Top             =   2835
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   61603843
         CurrentDate     =   39651
      End
      Begin MSComCtl2.DTPicker DTPProses 
         DataField       =   "tgl_packing"
         Height          =   315
         Index           =   4
         Left            =   4260
         TabIndex        =   34
         Tag             =   "LS"
         Top             =   3180
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   61603843
         CurrentDate     =   39651
      End
      Begin MSComCtl2.DTPicker DTPSample 
         DataField       =   "tgl_sample"
         Height          =   315
         Left            =   1875
         TabIndex        =   45
         Tag             =   "LS"
         Top             =   3900
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   61603843
         CurrentDate     =   39651
      End
      Begin MSComCtl2.DTPicker DTPKirim 
         DataField       =   "tgl_kirim1"
         Height          =   315
         Index           =   0
         Left            =   1875
         TabIndex        =   50
         Tag             =   "LS"
         Top             =   4725
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   61603843
         CurrentDate     =   39651
      End
      Begin MSComCtl2.DTPicker DTPKirim 
         DataField       =   "tgl_kirim2"
         Height          =   315
         Index           =   1
         Left            =   1875
         TabIndex        =   51
         Tag             =   "LS"
         Top             =   5055
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   61603843
         CurrentDate     =   39651
      End
      Begin MSComCtl2.DTPicker DTPKirim 
         DataField       =   "tgl_kirim3"
         Height          =   315
         Index           =   2
         Left            =   1875
         TabIndex        =   52
         Tag             =   "LS"
         Top             =   5385
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   61603843
         CurrentDate     =   39651
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "dateTrans"
         Height          =   345
         Left            =   1650
         TabIndex        =   65
         Tag             =   "LS"
         Top             =   555
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   61603843
         CurrentDate     =   39651
      End
      Begin VB.Line Line3 
         Index           =   2
         X1              =   2355
         X2              =   300
         Y1              =   4545
         Y2              =   4545
      End
      Begin VB.Line Line3 
         Index           =   1
         X1              =   2355
         X2              =   300
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         Height          =   285
         Index           =   29
         Left            =   285
         TabIndex        =   63
         Top             =   5850
         Width           =   1605
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         Height          =   285
         Index           =   28
         Left            =   6225
         TabIndex        =   61
         Top             =   5415
         Width           =   1605
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         Height          =   285
         Index           =   27
         Left            =   6225
         TabIndex        =   60
         Top             =   5085
         Width           =   1605
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         Height          =   285
         Index           =   26
         Left            =   6225
         TabIndex        =   59
         Top             =   4770
         Width           =   1605
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Berat "
         Height          =   285
         Index           =   25
         Left            =   4320
         TabIndex        =   55
         Top             =   5415
         Width           =   1605
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Berat "
         Height          =   285
         Index           =   24
         Left            =   4320
         TabIndex        =   54
         Top             =   5085
         Width           =   1605
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Berat "
         Height          =   285
         Index           =   23
         Left            =   4320
         TabIndex        =   53
         Top             =   4785
         Width           =   1605
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   2370
         X2              =   315
         Y1              =   1230
         Y2              =   1230
      End
      Begin VB.Line Line2 
         X1              =   2280
         X2              =   315
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2475
         X2              =   315
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Kirim "
         Height          =   285
         Index           =   22
         Left            =   315
         TabIndex        =   49
         Top             =   4770
         Width           =   1605
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Berat Proses"
         Height          =   285
         Index           =   21
         Left            =   315
         TabIndex        =   47
         Top             =   4275
         Width           =   1605
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Sample Lab"
         Height          =   285
         Index           =   20
         Left            =   315
         TabIndex        =   46
         Top             =   3945
         Width           =   1605
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Hari"
         Height          =   285
         Index           =   19
         Left            =   7890
         TabIndex        =   44
         Top             =   3225
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Hari"
         Height          =   285
         Index           =   18
         Left            =   7890
         TabIndex        =   43
         Top             =   2910
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Hari"
         Height          =   285
         Index           =   17
         Left            =   7890
         TabIndex        =   42
         Top             =   2550
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Hari"
         Height          =   285
         Index           =   16
         Left            =   7890
         TabIndex        =   41
         Top             =   2220
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Hari"
         Height          =   285
         Index           =   15
         Left            =   7890
         TabIndex        =   40
         Top             =   1875
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   285
         Index           =   14
         Left            =   3510
         TabIndex        =   29
         Top             =   3225
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   285
         Index           =   13
         Left            =   3510
         TabIndex        =   28
         Top             =   2880
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   285
         Index           =   12
         Left            =   3510
         TabIndex        =   27
         Top             =   2535
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   285
         Index           =   11
         Left            =   3510
         TabIndex        =   26
         Top             =   2190
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   285
         Index           =   10
         Left            =   3510
         TabIndex        =   25
         Top             =   1845
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Orang"
         Height          =   285
         Index           =   9
         Left            =   2730
         TabIndex        =   24
         Top             =   3225
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Orang"
         Height          =   285
         Index           =   7
         Left            =   2730
         TabIndex        =   23
         Top             =   2880
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Orang"
         Height          =   285
         Index           =   6
         Left            =   2730
         TabIndex        =   22
         Top             =   2535
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Orang"
         Height          =   285
         Index           =   5
         Left            =   2730
         TabIndex        =   21
         Top             =   2190
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   285
         Index           =   8
         Left            =   3090
         TabIndex        =   20
         Top             =   4020
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Orang"
         Height          =   285
         Index           =   3
         Left            =   2730
         TabIndex        =   19
         Top             =   1845
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PROSES"
         Height          =   285
         Index           =   4
         Left            =   270
         TabIndex        =   5
         Top             =   1485
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Berat Datang"
         Height          =   285
         Index           =   2
         Left            =   315
         TabIndex        =   4
         Top             =   990
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Datang"
         Height          =   285
         Index           =   1
         Left            =   315
         TabIndex        =   3
         Top             =   630
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   285
         Index           =   0
         Left            =   315
         TabIndex        =   2
         Top             =   270
         Width           =   1305
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   6915
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   1005
      BindFormTAG     =   "LS"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmLembarSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RsCall As New DBQuick

Private Sub cmdLink_Click(Index As Integer)
   OpenPartner Index
End Sub

Private Sub Form_Load()
   HiasFormManTell Picture1, Me
   MyDDE.SetPermissions = aksess.MayDo("Lembar Supplier") 'set hak aksess
   Set MyDDE.BindForm = Me
   Set MyDDE.ActiveConnection = CNN
   MyDDE.PrepareQuery = "select * from view_lembar_supplier"
   Set mCall = New frmCaller
End Sub


Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
   MyDDE.GetFieldByName("partnerID") = mCall.GetFieldByName("kode Supplier")
   MyDDE.GetFieldByName("TransID") = mCall.GetFieldByName("No Penerimaan")
   MyDDE.GetFieldByName("CompanyName") = mCall.GetFieldByName("Nama supplier")
   MyDDE.GetFieldByName("qty_receive") = mCall.GetFieldByName("berat")
   MyDDE.GetFieldByName("kondisi") = mCall.GetFieldByName("kondisi")
   If mCall.GetFieldByName("kondisi") = "Basah" Then
      Op(0).Value = True
   Else
      Op(0).Value = False
   End If
End Sub


Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
   Dim x As Integer
   cmdLink(0).Enabled = False
   Select Case AdReasonActiveDb
      Case tmbAddNew
         cmdLink(0).Enabled = True
         DTPicker1.Value = Now
         DTPSample.Value = Now
         For x = 0 To 4
            DTPProses(x).Value = Now
            chkProses(x).Value = 0
         Next
         
         For x = 0 To 2
            DTPKirim(x).Value = Now
         Next
         
         
   End Select
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   If MyDDE.GetFieldByName("kondisi") = "Basah" Then
      Op(0).Value = True
      Op(1).Value = False
   Else
      Op(0).Value = False
      Op(1).Value = True
   End If
End Sub


Private Sub PrepareSQL()
   MyDDE.PrepareAppend = "INSERT INTO [lembar_supplier]([Partnerid],[no_penerimaan] " & _
                                 ",[cuci],[jemur],[sortir],[napel],[packing],[org_cuci],[org_jemur],[org_sortir],[org_napel] " & _
                                 ",[org_packing],[tgl_cuci],[tgl_jemur],[tgl_sortir],[tgl_napel],[tgl_packing],[hari_cuci] " & _
                                 ",[hari_jemur],[hari_sortir],[hari_napel],[hari_packing],[tgl_sample],[berat_proses],[tgl_kirim1] " & _
                                 ",[tgl_kirim2],[tgl_kirim3],[berat_kirim1],[berat_kirim2],[berat_kirim3],[keterangan]) " & _
                          " Values ('" & MyDDE.GetFieldByName("partnerID") & "','" & MyDDE.GetFieldByName("TransID") & "'" & _
                                 "," & chkProses(0).Value & _
                                 "," & chkProses(1).Value & _
                                 "," & chkProses(2).Value & _
                                 "," & chkProses(3).Value & _
                                 "," & chkProses(4).Value & _
                                 "," & FQty(txtJmlOrang(0)) & "," & FQty(txtJmlOrang(1)) & _
                                 "," & FQty(txtJmlOrang(2)) & "," & FQty(txtJmlOrang(3)) & _
                                 "," & FQty(txtJmlOrang(4)) & ",'" & Format(DTPProses(0).Value, "yyyy-MM-dd") & "'" & _
                                 ",'" & Format(DTPProses(1).Value, "yyyy-MM-dd") & "'" & _
                                 ",'" & Format(DTPProses(2).Value, "yyyy-MM-dd") & "'" & _
                                 ",'" & Format(DTPProses(3).Value, "yyyy-MM-dd") & "'" & _
                                 ",'" & Format(DTPProses(4).Value, "yyyy-MM-dd") & "'" & _
                                 ", " & FQty(txtHari(0)) & ", " & FQty(txtHari(1)) & _
                                 ", " & FQty(txtHari(2)) & ", " & FQty(txtHari(3)) & _
                                 ", " & FQty(txtHari(4)) & ",'" & Format(DTPSample.Value, "yyyy-MM-dd") & "'" & _
                                 ", " & txtBeratProses & ",'" & Format(DTPKirim(0), "yyyy-mm-dd") & "'" & _
                                 ",'" & Format(DTPKirim(1), "yyyy-mm-dd") & "'" & _
                                 ",'" & Format(DTPKirim(2), "yyyy-mm-dd") & "'" & _
                                 ", " & FQty(txtBeratKirim(0)) & "," & FQty(txtBeratKirim(1)) & "," & FQty(txtBeratKirim(2)) & ",'" & txtKet & "')"


   MyDDE.PrepareUpdate = "UPDATE [lembar_supplier] set [Partnerid] ='" & MyDDE.GetFieldByName("partnerID") & "'" & _
                                    ",[no_penerimaan] ='" & MyDDE.GetFieldByName("TRansID") & "'" & _
                                    ",[cuci] = " & chkProses(0).Value & _
                                    ",[jemur] = " & chkProses(1).Value & _
                                    ",[sortir] = " & chkProses(2).Value & _
                                    ",[napel] = " & chkProses(3).Value & _
                                    ",[packing] =" & chkProses(4).Value & _
                                    ",[org_cuci] =" & FQty(txtJmlOrang(0)) & ",[org_jemur]=" & FQty(txtJmlOrang(1)) & _
                                    ",[org_sortir]=" & FQty(txtJmlOrang(2)) & ",[org_napel] =" & FQty(txtJmlOrang(3)) & _
                                    ",[org_packing]=" & FQty(txtJmlOrang(4)) & ",[tgl_cuci]='" & Format(DTPProses(0).Value, "yyyy-MM-dd") & "'" & _
                                    ",[tgl_jemur]='" & Format(DTPProses(1).Value, "yyyy-MM-dd") & "'" & _
                                    ",[tgl_sortir]='" & Format(DTPProses(2).Value, "yyyy-MM-dd") & "'" & _
                                    ",[tgl_napel]='" & Format(DTPProses(3).Value, "yyyy-MM-dd") & "'" & _
                                    ",[tgl_packing]='" & Format(DTPProses(4).Value, "yyyy-MM-dd") & "'" & _
                                    ",[hari_cuci] =" & FQty(txtHari(0)) & ",[hari_jemur] =" & FQty(txtHari(1)) & _
                                    ",[hari_sortir] =" & FQty(txtHari(2)) & ",[hari_napel] =" & FQty(txtHari(3)) & _
                                    ",[hari_packing]=" & FQty(txtHari(4)) & ",[tgl_sample]='" & Format(DTPSample.Value, "yyyy-MM-dd") & "'" & _
                                    ",[berat_proses]=" & txtBeratProses & _
                                    ",[tgl_kirim1] ='" & Format(DTPKirim(0), "yyyy-mm-dd") & "'" & _
                                    ",[tgl_kirim2] ='" & Format(DTPKirim(1), "yyyy-mm-dd") & "'" & _
                                    ",[tgl_kirim3] ='" & Format(DTPKirim(2), "yyyy-mm-dd") & "'" & _
                                    ",[berat_kirim1] =" & FQty(txtBeratKirim(0)) & _
                                    ",[berat_kirim2] =" & FQty(txtBeratKirim(1)) & _
                                    ",[berat_kirim3] =" & FQty(txtBeratKirim(2)) & ",[keterangan]='" & txtKet & "' " & _
                           "Where idx='" & MyDDE.GetFieldByName("idx") & "'"


   MyDDE.PrepareDelete = "delete from [lembar_supplier] where idx='" & MyDDE.GetFieldByName("idx") & "'"
   
End Sub


Private Sub OpenPartner(Index As Integer)
   Select Case Index
          Case 0: RsCall.DBOpen "SELECT * from view_lembar_supplier_get_supplier ", CNN, lckLockReadOnly
   End Select
   If RsCall.Recordcount <> 0 Then
       Select Case Index
              Case 0: mCall.FromTagActive = "Supplier"
       End Select
       Set mCall.FormData = RsCall.DBRecordset
       mCall.LookUp Me
   Else
      MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgCrtical
   End If
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

   Select Case AdReasonActiveDb
      Case tmbSave:
         If MyDDE.CheckEmptyControl = False Then
            MyDDE.IsChildMemberReady = True
            PrepareSQL
         Else
            MyDDE.IsChildMemberReady = False
         End If
      Case tmbDelete:
         PrepareSQL
   End Select

End Sub
