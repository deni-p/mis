VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmOutstandingPO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Outstanding Pembelian"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10635
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOutstandingPO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   10635
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
      ScaleWidth      =   10635
      TabIndex        =   3
      Top             =   5700
      Width           =   10635
      Begin VB.CommandButton cmd 
         Caption         =   "&Detil"
         Enabled         =   0   'False
         Height          =   555
         Index           =   6
         Left            =   75
         Picture         =   "frmOutstandingPO.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "F9 Tambah Detail Transaksi"
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
         Left            =   -75
         TabIndex        =   5
         Top             =   0
         Width           =   10965
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Close Penerimaan"
         Height          =   555
         Index           =   2
         Left            =   795
         Picture         =   "frmOutstandingPO.frx":D0A4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   60
         Width           =   1530
      End
      Begin VB.CommandButton cmd 
         Caption         =   "E&xit"
         Height          =   555
         Index           =   0
         Left            =   2325
         Picture         =   "frmOutstandingPO.frx":138F6
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
      ScaleWidth      =   10635
      TabIndex        =   4
      Top             =   0
      Width           =   10635
      Begin MSDataGridLib.DataGrid grid 
         Height          =   5535
         Left            =   90
         TabIndex        =   0
         Top             =   75
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   9763
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "POID"
            Caption         =   "No PO"
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
            DataField       =   "Supplier"
            Caption         =   "Supplier"
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
            DataField       =   "Tanggal PO"
            Caption         =   "Tanggal PO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd MMM yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Tgl Kebutuhan"
            Caption         =   "Tgl Kebutuhan"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd MMM yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "no barang"
            Caption         =   "No Barang"
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
            DataField       =   "Nama Barang"
            Caption         =   "Nama Barang"
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
            DataField       =   "Qty PO"
            Caption         =   "Qty PO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "QTy Diterima"
            Caption         =   "Qty Diterima"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "Harga"
            Caption         =   "Harga"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmOutstandingPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsOut As New DBQuick

Private Sub cmd_Click(Index As Integer)
   Select Case Index
      Case 0: Unload Me
      Case 1:
         If rsOut.Recordcount > 0 Then
            FrmPurchasing.IDParams = rsOut.DBRecordset.Fields("POID")
            FrmPurchasing.SetFocus
         Else
            MessageBox "Tidak Ada data yang tersedia", "Stop", msgOkOnly, msgExclamation
         End If
      Case 2:
         If MessageBox("Yakin Data ini Akan Ditutup ?", "Konfirmasi", msgYesNo, msgQuestion) = 1 Then
            If rsOut.Fields("statusTrans") = 2 Then
               SendDataToServer "update [detail PO] set statusTrans = 8 where PurchaseID='" & rsOut.Fields("POID") & "' and noItem='" & rsOut.Fields("No Barang") & "'"
               rsOut.DBOpen "select * from OutstandingPO order by [Tanggal PO]", CNN
               Set grid.DataSource = rsOut.DBRecordset
            Else
               MessageBox "Data tidak bisa diproses karena barang belum diterima sama sekali", "Peringatan", msgOkOnly, msgCrtical
            End If
         End If
   End Select
End Sub

Private Sub Form_Load()
   HiasFormManTell Picture2, Me
   rsOut.DBOpen "select * from OutstandingPO order by [Tanggal PO]", CNN
   grid.HeadLines = 2
   Set grid.DataSource = rsOut.DBRecordset
End Sub

