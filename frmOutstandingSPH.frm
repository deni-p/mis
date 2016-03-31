VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmOutstandingSPH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Outstanding SPH "
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10725
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOutstandingSPH.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   10725
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0FFFF&
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
      Height          =   690
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   10725
      TabIndex        =   3
      Top             =   5700
      Width           =   10725
      Begin VB.Frame FrTombol 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   30
         Left            =   -45
         TabIndex        =   6
         Top             =   0
         Width           =   10995
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Close SPH"
         Height          =   555
         Index           =   2
         Left            =   60
         Picture         =   "frmOutstandingSPH.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   60
         Width           =   945
      End
      Begin VB.CommandButton cmd 
         Caption         =   "E&xit"
         Height          =   555
         Index           =   0
         Left            =   1005
         Picture         =   "frmOutstandingSPH.frx":D0A4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   60
         Width           =   720
      End
   End
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
      Height          =   5775
      Left            =   0
      ScaleHeight     =   5775
      ScaleWidth      =   10725
      TabIndex        =   4
      Top             =   0
      Width           =   10725
      Begin MSDataGridLib.DataGrid grid 
         Height          =   2550
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   4498
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   15
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            MarqueeStyle    =   3
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid grid 
         Height          =   2625
         Index           =   1
         Left            =   105
         TabIndex        =   5
         Top             =   2925
         Width           =   10530
         _ExtentX        =   18574
         _ExtentY        =   4630
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   15
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            MarqueeStyle    =   3
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmOutstandingSPH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsOut As New DBQuick
Dim RsDetail As New DBQuick

Private Sub cmd_Click(Index As Integer)
   Select Case Index
      Case 0: Unload Me
      Case 2:
         If MessageBox("Yakin Data ini Akan Ditutup ?", "Konfirmasi", msgYesNo, msgQuestion) = 1 Then
               SendDataToServer "update spph_header set status = 1 where SPPHID='" & rsOut.Fields("SPPHID") & "'"
               rsOut.DBOpen "select * from OutstandingSPH_Header order by [Tgl Transaksi] desc", CNN
               Set grid(0).DataSource = rsOut.DBRecordset
         End If
   End Select
End Sub

Private Sub Form_Load()
   HiasFormManTell Picture2, Me
   rsOut.DBOpen "select * from OutstandingSPH_Header order by [Tgl Transaksi] desc", CNN
   grid(0).HeadLines = 2
   grid(1).HeadLines = 2
   
   Set grid(0).DataSource = rsOut.DBRecordset
   
   If rsOut.DBRecordset.Recordcount > 0 Then
      loadDetail
   End If
End Sub


Private Sub loadDetail()
   RsDetail.DBOpen "select * from OutstandingSPH_line where SPPHID ='" & rsOut.DBRecordset.Fields("SPPHID") & "'", CNN
   Set grid(1).DataSource = RsDetail.DBRecordset
End Sub

Private Sub grid_Click(Index As Integer)
   If Index = 0 Then loadDetail
End Sub

Private Sub grid_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
   If Index = 0 Then loadDetail
End Sub

