VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FRMMIXING_CHIP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MIXING CHIPS"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMMIXING_CHIP.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   10095
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
      Height          =   3405
      Left            =   0
      ScaleHeight     =   3405
      ScaleWidth      =   10095
      TabIndex        =   8
      Top             =   0
      Width           =   10095
      Begin MSDataGridLib.DataGrid GridDetail 
         Height          =   2700
         Left            =   5670
         TabIndex        =   6
         Top             =   180
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   4763
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   1
         BeginProperty Column00 
            DataField       =   "noEkstraksi"
            Caption         =   "No Ekstraksi"
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
               Locked          =   -1  'True
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "grup"
         DataSource      =   "DDE"
         Height          =   345
         Index           =   2
         Left            =   2625
         TabIndex        =   1
         Tag             =   "mixing"
         Top             =   960
         Width           =   705
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "prelot"
         DataSource      =   "DDE"
         Height          =   345
         Index           =   1
         Left            =   2625
         TabIndex        =   0
         Tag             =   "mixing"
         Top             =   570
         Width           =   1740
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "total_sebelum"
         DataSource      =   "DDE"
         Height          =   345
         Index           =   3
         Left            =   2625
         TabIndex        =   2
         Tag             =   "mixing"
         Top             =   1335
         Width           =   705
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "total_sesudah"
         DataSource      =   "DDE"
         Height          =   345
         Index           =   4
         Left            =   2625
         TabIndex        =   3
         Tag             =   "mixing"
         Top             =   1710
         Width           =   705
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "tanggal_mulai"
         DataSource      =   "DDE"
         Height          =   345
         Index           =   0
         Left            =   2625
         TabIndex        =   4
         Tag             =   "mixing"
         Top             =   2085
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   609
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
         CustomFormat    =   "dd MMM yyyy hh:mm:ss"
         Format          =   20971523
         CurrentDate     =   39426
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "tanggal_selesai"
         DataSource      =   "DDE"
         Height          =   345
         Index           =   1
         Left            =   2625
         TabIndex        =   5
         Tag             =   "mixing"
         Top             =   2460
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   609
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
         CustomFormat    =   "dd MMM yyyy hh:mm:ss"
         Format          =   20971523
         CurrentDate     =   39426
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   2700
         X2              =   240
         Y1              =   1290
         Y2              =   1290
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2670
         X2              =   240
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Grup"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   1035
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pre Lot Chips No"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   660
         Width           =   1290
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   2730
         X2              =   240
         Y1              =   1665
         Y2              =   1665
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Chips sesudah diblender                       Kg"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   11
         Top             =   1785
         Width           =   3855
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   2730
         X2              =   240
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   2730
         X2              =   240
         Y1              =   2415
         Y2              =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal dan waktu mulai"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   10
         Top             =   2175
         Width           =   1920
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   2730
         X2              =   240
         Y1              =   2790
         Y2              =   2790
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal dan waktu selesai"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   9
         Top             =   2535
         Width           =   1920
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Chips sebelum diblender                       Kg"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   12
         Top             =   1410
         Width           =   3360
      End
   End
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   7
      Top             =   3420
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   1005
      BindFormTAG     =   "mixing"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FRMMIXING_CHIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Dim rsno As New DBQuick
Private rsDetail As New DBQuick
Private IDGen As New IDGenerator


Private Sub LoadEkstraksi()
rsno.DBOpen "select * from statusProduksi where status=1 and posisi='CRUSHER'", CNN
   Set mCall = New frmCaller
   Set mCall.FormData = rsno.DBRecordset
   mCall.FromTagActive = "Extraction No"
   mCall.CaptionLink = "Extraction No"
End Sub

Private Sub DDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
   Case tmbAddNew:
        DDE.GetFieldByName("prelot") = IDGen.GetID("PL")
        tgl(0).Value = Now
        tgl(1).Value = Now
        DDE.GetFieldByName("tanggal_mulai") = Now
        DDE.GetFieldByName("tanggal_selesai") = Now
   Case tmbDetail:
      LoadEkstraksi
   Case tmbSave:
      If DDE.IsChildMemberReady = True Then
         SendDataToServer "insert into statusProduksi(noEkstraksi,posisi,status,tanggal) values ('" & _
                           txt(1).Text & "','MIXING',1,'" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "')"
         SimpanDetail
      End If
End Select
End Sub

Private Sub SimpanDetail()
   If SendDataToServer("delete from mixing_detail where prelot ='" & txt(1).Text & "'") = True Then
      With DDE.ChildRecordset
         .MoveFirst
         While Not .EOF
            SendDataToServer "insert into mixing_detail (prelot,NoEkstraksi) values ('" & txt(1).Text & "','" & .Fields("NoEkstraksi") & "')"
            .MoveNext
         Wend
      End With
   End If
End Sub

Private Sub DDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   rsDetail.DBOpen "select noEkstraksi from mixing_detail where prelot ='" & txt(1).Text & "'", CNN
   Set DDE.ChildRecordset = rsDetail.DBRecordset
   Set GridDetail.DataSource = DDE.ChildRecordset
End Sub

Private Sub DDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
Case tmbSave:
   DDE.IsChildMemberReady = True
   simpan
Case tmbDelete
   simpan
End Select
End Sub

Function simpan()
   DDE.PrepareAppend = "insert into MIXING_header (prelot,grup,total_sebelum,total_sesudah,tanggal_mulai,tanggal_selesai) values ('" & _
                              txt(1).Text & "', '" & txt(2).Text & "', " & FQty(txt(3).Text) & ", " & FQty(txt(4).Text) & " , '" & Format(tgl(0).Value, "yyyy-MM-dd hh:mm:ss") & "', '" & Format(tgl(1).Value, "yyyy-MM-dd hh:mm:ss") & "')"
   DDE.PrepareUpdate = "update MIXING_header set , grup = '" & txt(2).Text & "', total_sebelum = " & FQty(txt(3).Text) & ", total_sesudah = " & FQty(txt(4).Text) & ", tanggal_mulai = '" & Format(tgl(0).Value, "yyyy-MM-dd hh:mm:ss") & "', tanggal_selesai = '" & Format(tgl(1).Value, "yyyy-MM-dd hh:mm:ss") & "' where prelot = '" & txt(1).Text & "'"
   DDE.PrepareDelete = "delete from  mixing_header where prelot ='" & txt(1).Text & "'"
End Function

Private Sub Form_Load()
With DDE
Set .BindForm = Me
    .BindFormTAG = "mixing"
Set .ActiveConnection = CNN
    .PrepareQuery = "select * from mixing_header"
End With
HiasForm Picture2, Me
End Sub


Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
   DDE.ChildRecordset.Fields("NoEkstraksi") = mCall.GetFieldByName("noEkstraksi")
End Sub

