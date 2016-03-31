VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmApproval 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Approval"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmApproval.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   11295
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   11295
      TabIndex        =   4
      Top             =   7050
      Width           =   11295
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh"
         Height          =   555
         Left            =   825
         Picture         =   "FrmApproval.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Set Ke Surat Permintaan Penawaran Harga"
         Top             =   60
         Width           =   720
      End
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
         Width           =   11400
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Valid"
         Height          =   555
         Left            =   105
         Picture         =   "FrmApproval.frx":D0A4
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Set Ke Surat Permintaan Penawaran Harga"
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "K&eluar"
         Height          =   555
         Left            =   10485
         Picture         =   "FrmApproval.frx":138F6
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   75
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   7710
      Left            =   -60
      ScaleHeight     =   7710
      ScaleWidth      =   11460
      TabIndex        =   5
      Top             =   -45
      Width           =   11460
      Begin MSDataGridLib.DataGrid gridHeader 
         Height          =   4440
         Left            =   180
         TabIndex        =   0
         Top             =   150
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   7832
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
      Begin MSDataGridLib.DataGrid gridDetail 
         Height          =   2160
         Left            =   180
         TabIndex        =   1
         Top             =   4770
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   3810
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
            Weight          =   700
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
Attribute VB_Name = "frmApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private aValidasi As clsApproval

Public Property Let Validasi(aData As clsApproval)
   Set aValidasi = aData
   aValidasi.Start
   Set gridHeader.DataSource = aValidasi.MasterRecordset.DBRecordset
   Set gridDetail.DataSource = aValidasi.DetailRecordset.DBRecordset
   gridHeader.Columns(0).Alignment = dbgCenter
   gridHeader.Columns(0).width = 1100
End Property

Private Sub cmd_Click()
   Dim X As Integer
   If aValidasi.MasterRecordset.DBRecordset.Recordcount > 0 Then
      X = aValidasi.MasterRecordset.DBRecordset.Bookmark
      aValidasi.Approve
      aValidasi.Start
      Set gridHeader.DataSource = aValidasi.MasterRecordset.DBRecordset
      Set gridDetail.DataSource = aValidasi.DetailRecordset.DBRecordset
      gridHeader.Columns(0).Alignment = dbgCenter
      gridHeader.Columns(0).width = 1100
      On Error Resume Next
      aValidasi.MasterRecordset.DBRecordset.Move X - 1
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   Dim X As Integer
   X = aValidasi.MasterRecordset.DBRecordset.Bookmark
   aValidasi.Start
   Set gridHeader.DataSource = aValidasi.MasterRecordset.DBRecordset
   Set gridDetail.DataSource = aValidasi.DetailRecordset.DBRecordset
   gridHeader.Columns(0).Alignment = dbgCenter
   gridHeader.Columns(0).width = 1100
   On Error Resume Next
   aValidasi.MasterRecordset.DBRecordset.Move X - 1
End Sub

Private Sub Form_Load()
   HiasFormManTell Picture1, Me
   gridHeader.HeadLines = 2
   gridDetail.HeadLines = 2
End Sub

Private Sub gridHeader_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   Set gridDetail.DataSource = aValidasi.DetailRecordset.DBRecordset
End Sub

