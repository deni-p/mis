VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D7BB8F75-AC9E-4E80-A526-70EA20ACFD16}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form SHIEVER_POWDER 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SHIEVER POWDER PRE LOT NO"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9150
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   9150
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      ForeColor       =   &H80000008&
      Height          =   3165
      Left            =   -15
      ScaleHeight     =   3135
      ScaleWidth      =   9165
      TabIndex        =   1
      Top             =   -15
      Width           =   9195
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "pre_lot_powder"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   2445
         TabIndex        =   7
         Tag             =   "powder"
         Top             =   630
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "grup"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   2445
         TabIndex        =   6
         Tag             =   "powder"
         Top             =   900
         Width           =   1005
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "total_1"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   2430
         TabIndex        =   5
         Tag             =   "powder"
         Top             =   1800
         Width           =   645
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4170
         MaskColor       =   &H000000C0&
         Picture         =   "SHIEVER_POWDER.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "powder"
         Top             =   630
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "total_2"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2430
         TabIndex        =   3
         Tag             =   "powder"
         Top             =   2070
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "total"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2430
         TabIndex        =   2
         Tag             =   "powder"
         Top             =   2340
         Width           =   645
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "tgl_mulai"
         DataSource      =   "DDE"
         Height          =   345
         Index           =   2
         Left            =   2445
         TabIndex        =   8
         Tag             =   "powder"
         Top             =   1155
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MMMM-dd hh:mm:ss"
         Format          =   17694723
         CurrentDate     =   39423
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "tgl_selesai"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   3
         Left            =   2430
         TabIndex        =   9
         Tag             =   "powder"
         Top             =   1485
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MMMM-dd hh:mm:ss"
         Format          =   17694723
         CurrentDate     =   39423
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pre Lot Powder No"
         Height          =   255
         Index           =   13
         Left            =   420
         TabIndex        =   16
         Top             =   645
         Width           =   1755
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal dan waktu mulai"
         Height          =   255
         Index           =   14
         Left            =   420
         TabIndex        =   15
         Top             =   1245
         Width           =   2490
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal dan waktu selesai"
         Height          =   255
         Index           =   15
         Left            =   420
         TabIndex        =   14
         Top             =   1530
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Powder mesh > 150                    Kg"
         Height          =   255
         Index           =   16
         Left            =   405
         TabIndex        =   13
         Top             =   1845
         Width           =   3000
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Powder mesh  150                      Kg"
         Height          =   255
         Index           =   17
         Left            =   405
         TabIndex        =   12
         Top             =   2130
         Width           =   3045
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   420
         X2              =   2760
         Y1              =   1785
         Y2              =   1785
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   420
         X2              =   2760
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Line Line2 
         Index           =   2
         X1              =   420
         X2              =   2760
         Y1              =   2070
         Y2              =   2070
      End
      Begin VB.Line Line2 
         Index           =   3
         X1              =   420
         X2              =   2520
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Line Line2 
         Index           =   4
         X1              =   420
         X2              =   2745
         Y1              =   2610
         Y2              =   2610
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Grup"
         Height          =   255
         Index           =   0
         Left            =   420
         TabIndex        =   11
         Top             =   960
         Width           =   465
      End
      Begin VB.Line Line2 
         Index           =   5
         X1              =   435
         X2              =   2775
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Line Line2 
         Index           =   6
         X1              =   435
         X2              =   2775
         Y1              =   1485
         Y2              =   1485
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Powder                                       Kg"
         Height          =   255
         Index           =   1
         Left            =   405
         TabIndex        =   10
         Top             =   2385
         Width           =   2940
      End
   End
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   3180
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   1005
      BindFormTAG     =   "powder"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "SHIEVER_POWDER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents shiever As frmCaller
Attribute shiever.VB_VarHelpID = -1
Dim rspowder As New DBQuick
Dim tabel As String

Private Sub cmdLink_Click()
    rspowder.DBOpen "select * from mixing_chips", CNN
    Set shiever = New frmCaller
    Set shiever.FormData = rspowder.DBRecordset
    shiever.FromTagActive = "MIXING CHIPS"
    shiever.CaptionLink = "MIXING CHIPS"
End Sub

Private Sub DDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb

        Case tmbAddNew:
            CmdLink.Enabled = True
            Text1(16).Enabled = False
            DTPicker1(2).Enabled = True
            DTPicker1(3).Enabled = True

        Case tmbEdit
            DTPicker1(2).Enabled = True
            DTPicker1(3).Enabled = True
            Text1(16).Enabled = False
    End Select

End Sub

Private Sub DDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb

        Case tmbSave:
            DDE.IsChildMemberReady = True
            simpan

        Case tmbDelete:
            DDE.PrepareDelete = "delete " & tabel & " where pre_lot_powder = '" & DDE.GetFieldByName("pre_lot_powder") & "'"

    End Select

End Sub

Function simpan()
    DDE.PrepareAppend = " insert into " & tabel & " (PRE_LOT_POWDER, GRUP, TGL_MULAI, TGL_SELESAI, TOTAL_1, TOTAL_2, TOTAL) values ('" & Text1(16).Text & "', '" & Text1(11).Text & "', '" & Format(DTPicker1(2).value, "yyyy-MM-dd hh:mm:ss") & "', '" & Format(DTPicker1(3).value, "yyyy-MM-dd hh:mm:ss") & "', '" & DDE.GetFieldByName("total_1") & "', '" & DDE.GetFieldByName("total_2") & "', '" & DDE.GetFieldByName("total") & "')"
    DDE.PrepareUpdate = " update " & tabel & " set grup = '" & Text1(11).Text & "', tgl_mulai = '" & Format(DTPicker1(2).value, "yyyy-MM-dd hh:mm:ss") & "', tgl_selesai = '" & Format(DTPicker1(3).value, "yyyy-MM-dd hh:mm:ss") & "' total_1 = '" & DDE.GetFieldByName("total_1") & "', total_2 = '" & DDE.GetFieldByName("total_2") & "', total = '" & DDE.GetFieldByName("total") & "'    "
End Function

Private Sub Form_Load()
    On Error Resume Next
    Dim HuruF As Control

    With DDE
        Set .BindForm = Me
        .BindFormTAG = "powder"
        Set .ActiveConnection = CNN
        .PrepareQuery = "select * from SHIEVER_POWDER_PRE_LOT_NO"
    End With

    HiasForm Picture2, Me
    seting Me
    tabel = "SHIEVER_POWDER_PRE_LOT_NO"
End Sub

Private Sub shiever_RowColChange(ByVal TagForm As String, _
                                 ByVal pRecordset As ADODB.Recordset)
    Text1(16).Text = rspowder.DBRecordset.Fields("pre_lot_chip_no")
End Sub
