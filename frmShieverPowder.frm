VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B1E614FF-F86D-4F68-A86F-2584A0570C66}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmShieverPowder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SHIEVER POWDER PRE LOT NO"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShieverPowder.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   9150
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
      Height          =   3165
      Left            =   0
      ScaleHeight     =   3165
      ScaleWidth      =   9150
      TabIndex        =   1
      Top             =   0
      Width           =   9150
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "pre_lot_powder"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   315
         Index           =   16
         Left            =   2430
         TabIndex        =   7
         Tag             =   "powder"
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "grup"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   315
         Index           =   11
         Left            =   2430
         TabIndex        =   6
         Tag             =   "powder"
         Top             =   585
         Width           =   1905
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "total_1"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   315
         Index           =   12
         Left            =   2430
         TabIndex        =   5
         Tag             =   "powder"
         Top             =   1635
         Width           =   975
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5100
         MaskColor       =   &H000000C0&
         Picture         =   "frmShieverPowder.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "powder"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   390
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "total_2"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   2430
         TabIndex        =   3
         Tag             =   "powder"
         Top             =   1980
         Width           =   960
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "total"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   2430
         TabIndex        =   2
         Tag             =   "powder"
         Top             =   2325
         Width           =   960
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "tgl_mulai"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   2
         Left            =   2430
         TabIndex        =   8
         Tag             =   "powder"
         Top             =   930
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMM yyyy hh:mm:ss"
         Format          =   51183619
         CurrentDate     =   39423
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "tgl_selesai"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   3
         Left            =   2430
         TabIndex        =   9
         Tag             =   "powder"
         Top             =   1275
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMM yyyy hh:mm:ss"
         Format          =   51183619
         CurrentDate     =   39423
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   435
         X2              =   2775
         Y1              =   1575
         Y2              =   1575
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   435
         X2              =   2775
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Line2 
         Index           =   2
         X1              =   435
         X2              =   2775
         Y1              =   1935
         Y2              =   1935
      End
      Begin VB.Line Line2 
         Index           =   3
         X1              =   435
         X2              =   2535
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line2 
         Index           =   5
         X1              =   435
         X2              =   2775
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Line Line2 
         Index           =   6
         X1              =   435
         X2              =   2775
         Y1              =   1230
         Y2              =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pre Lot Powder No"
         Height          =   255
         Index           =   13
         Left            =   435
         TabIndex        =   16
         Top             =   270
         Width           =   1755
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal dan waktu mulai"
         Height          =   255
         Index           =   14
         Left            =   435
         TabIndex        =   15
         Top             =   975
         Width           =   2490
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal dan waktu selesai"
         Height          =   255
         Index           =   15
         Left            =   435
         TabIndex        =   14
         Top             =   1335
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Powder mesh > 150                           Kg"
         Height          =   255
         Index           =   16
         Left            =   435
         TabIndex        =   13
         Top             =   1710
         Width           =   4695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Powder mesh  150                              Kg"
         Height          =   255
         Index           =   17
         Left            =   435
         TabIndex        =   12
         Top             =   2040
         Width           =   4755
      End
      Begin VB.Line Line2 
         Index           =   4
         X1              =   435
         X2              =   2760
         Y1              =   2625
         Y2              =   2625
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Grup"
         Height          =   255
         Index           =   0
         Left            =   435
         TabIndex        =   11
         Top             =   630
         Width           =   465
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Powder                                               Kg"
         Height          =   255
         Index           =   1
         Left            =   435
         TabIndex        =   10
         Top             =   2385
         Width           =   3885
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
Attribute VB_Name = "frmShieverPowder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
     cmdLink.Enabled = True
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

Private Sub shiever_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Text1(16).Text = rspowder.DBRecordset.Fields("pre_lot_chip_no")
End Sub
