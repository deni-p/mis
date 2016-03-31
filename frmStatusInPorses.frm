VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStatusInPorses 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Status In Proses"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7875
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStatusInPorses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7875
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   7875
      TabIndex        =   5
      Top             =   5130
      Width           =   7875
      Begin VB.CommandButton CmdTombol 
         Caption         =   "&Confirm Finish"
         Height          =   555
         Index           =   1
         Left            =   45
         Picture         =   "frmStatusInPorses.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   45
         Width           =   1215
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "E&xit"
         Height          =   555
         Index           =   0
         Left            =   7095
         Picture         =   "frmStatusInPorses.frx":D0A4
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   45
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   5115
      Left            =   0
      ScaleHeight     =   5115
      ScaleWidth      =   7875
      TabIndex        =   4
      Top             =   0
      Width           =   7875
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EAAF6F&
         Caption         =   "   Pilih Semua"
         Height          =   330
         Left            =   105
         TabIndex        =   1
         Top             =   4740
         Width           =   2475
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4545
         Left            =   120
         TabIndex        =   0
         Top             =   135
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   8017
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No Ekstraksi"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Rekomendasi"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Maufacture Order"
            Object.Width           =   4410
         EndProperty
      End
   End
End
Attribute VB_Name = "frmStatusInPorses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RsLoad As New DBQuick



Private Sub Check1_Click()
   Dim x As Integer
   If Check1.Value = 1 Then
      For x = 1 To ListView1.ListItems.Count
         ListView1.ListItems(x).Checked = True
      Next
   End If
End Sub

Private Sub CmdTombol_Click(Index As Integer)
   Dim x As Integer
   Select Case Index
      Case 0: Unload Me
      Case 1:
         If MessageBox("Yakin Konfirmasi data akan dijalankan ? ", "Konfirmasi", msgYesNo) = 1 Then
            For x = 1 To ListView1.ListItems.Count
               If ListView1.ListItems(x).Checked Then
                  SendDataToServer "update [Manufacture Order] set status='FINISHED',FinishedDate='" & Format(Now, "yyyy-MM-dd") & "' where OrderID='" & ListView1.ListItems(x).SubItems(2) & "'"
               End If
            Next
            LoadData
         End If
   End Select
End Sub

Private Sub Form_Load()
   HiasFormManTell Picture1, Me
   LoadData
End Sub

Private Sub LoadData()
   Dim x As Integer
   RsLoad.DBOpen "SELECT NoEkstraksi as splNo, no_rekomendasi , [Manufacture Order].OrderID " & _
                 "FROM [Manufacture Order] inner JOIN " & _
                     " StatusProduksi ON [Manufacture Order].no_rekomendasi = StatusProduksi.rekomendasi " & _
                 " WHERE  (StatusProduksi.Posisi='CRUSHER') and ([Manufacture Order].status='RELEASED') ", CNN
   ListView1.ListItems.Clear
   
   If RsLoad.DBRecordset.Recordcount > 0 Then
      x = 1
      While Not RsLoad.DBRecordset.EOF
         ListView1.ListItems.Add x, "A" & RsLoad.DBRecordset.Fields(0), RsLoad.DBRecordset.Fields(0)
         ListView1.ListItems(x).SubItems(1) = IIf(IsNull(RsLoad.DBRecordset.Fields(1)), "", RsLoad.DBRecordset.Fields(1))
         ListView1.ListItems(x).SubItems(2) = IIf(IsNull(RsLoad.DBRecordset.Fields(2)), "", RsLoad.DBRecordset.Fields(2))
         RsLoad.DBRecordset.MoveNext
      Wend
   End If
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
   If Item.Checked = False Then Check1.Value = 0
End Sub
