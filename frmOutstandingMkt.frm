VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmOutstandingMkt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Outstanding"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10740
   Icon            =   "frmOutstandingMkt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   10740
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   10740
      TabIndex        =   3
      Top             =   6525
      Width           =   10740
      Begin VB.Frame FrTombol 
         BackColor       =   &H00C0FFFF&
         Height          =   30
         Left            =   -60
         TabIndex        =   10
         Top             =   0
         Width           =   10995
      End
      Begin VB.CommandButton cmd 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   0
         Left            =   1470
         Picture         =   "frmOutstandingMkt.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Close Contract"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   2
         Left            =   120
         Picture         =   "frmOutstandingMkt.frx":834C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         Width           =   1350
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   0
      ScaleHeight     =   6615
      ScaleWidth      =   10740
      TabIndex        =   0
      Top             =   0
      Width           =   10740
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Option"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   5055
         Begin VB.OptionButton Option3 
            BackColor       =   &H00EAAF6F&
            Caption         =   "Sales Order"
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
            Left            =   3360
            TabIndex        =   9
            Top             =   330
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00EAAF6F&
            Caption         =   "Sales Contract"
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
            Left            =   1680
            TabIndex        =   8
            Top             =   330
            Width           =   1530
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00EAAF6F&
            Caption         =   "Sales Quote"
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
            Left            =   120
            TabIndex        =   7
            Top             =   330
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin MSDataGridLib.DataGrid grid 
         Height          =   2550
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   1110
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
         TabIndex        =   2
         Top             =   3765
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
Attribute VB_Name = "frmOutstandingMkt"
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
               If Option1.Value = True Then
                    SendDataToServer "update [PO Detail] set status = 1 where purchaseid='" & rsOut.Fields("purchaseid") & "'"
                    rsOut.DBOpen "select * from [PO Order]  where typetrans='QUOTE' order by [datepurchase] ", CNN
                    Set grid(0).DataSource = rsOut.DBRecordset
               ElseIf Option2.Value = True Then
                    SendDataToServer "update [PO Detail] set status = 1 where purchaseid='" & rsOut.Fields("purchaseid") & "'"
                    rsOut.DBOpen "select * from [PO Order]  where typetrans='SC' order by [datepurchase] ", CNN
                    Set grid(0).DataSource = rsOut.DBRecordset
               Else
                    SendDataToServer "update [PO Detail] set status = 1 where purchaseid='" & rsOut.Fields("purchaseid") & "'"
                    rsOut.DBOpen "select * from [PO Order]  where typetrans='SO' order by [datepurchase] ", CNN
                    Set grid(0).DataSource = rsOut.DBRecordset
               End If
         End If
   End Select
End Sub

Private Sub Form_Load()
   HiasFormManTell Picture2, Me
End Sub


Private Sub loadDetail()
On Error Resume Next
   If Option1.Value = True Then
        RsDetail.DBOpen "select * from OutstandingSQ_Detail where purchaseid ='" & rsOut.DBRecordset.Fields("purchaseid") & "'", CNN
        Set grid(1).DataSource = RsDetail.DBRecordset
   ElseIf Option2.Value = True Then
        RsDetail.DBOpen "select * from OutstandingSC_Detail where purchaseid ='" & rsOut.DBRecordset.Fields("purchaseid") & "'", CNN
        Set grid(1).DataSource = RsDetail.DBRecordset
   Else
        RsDetail.DBOpen "select * from OutstandingSO_Detail where purchaseid ='" & rsOut.DBRecordset.Fields("purchaseid") & "'", CNN
        Set grid(1).DataSource = RsDetail.DBRecordset
   End If

End Sub

Private Sub grid_Click(Index As Integer)
   If Index = 0 Then loadDetail
End Sub

Private Sub grid_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
   If Index = 0 Then loadDetail
End Sub

Private Sub Option1_Click()
rsOut.DBOpen "select * from OutstandingQuote_Header order by [datepurchase] ", CNN
grid(0).HeadLines = 2
grid(1).HeadLines = 2

Set grid(0).DataSource = rsOut.DBRecordset

If rsOut.DBRecordset.Recordcount > 0 Then
   loadDetail
End If
cmd(2).Caption = "&Close " & "Quote"
End Sub

Private Sub Option2_Click()
rsOut.DBOpen "select * from OutstandingSC_Header order by [datepurchase] ", CNN
grid(0).HeadLines = 2
grid(1).HeadLines = 2
Set grid(0).DataSource = rsOut.DBRecordset
If rsOut.DBRecordset.Recordcount > 0 Then
   loadDetail
End If
cmd(2).Caption = "&Close " & "Contract"
End Sub

Private Sub Option3_Click()
rsOut.DBOpen "select * from OutstandingSO_Header order by [datepurchase] ", CNN
grid(0).HeadLines = 2
grid(1).HeadLines = 2
Set grid(0).DataSource = rsOut.DBRecordset
If rsOut.DBRecordset.Recordcount > 0 Then
   loadDetail
End If
cmd(2).Caption = "&Close " & "Order"
End Sub
