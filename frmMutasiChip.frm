VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMutasiChip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mutasi CHIP ke Gudang"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10860
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMutasiChip.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   10860
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   10860
      TabIndex        =   9
      Top             =   5310
      Width           =   10860
      Begin VB.CommandButton CmdTombol 
         Caption         =   "E&xit"
         Height          =   555
         Index           =   0
         Left            =   9915
         Picture         =   "frmMutasiChip.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   45
         Width           =   855
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "&Save"
         Height          =   555
         Index           =   1
         Left            =   75
         Picture         =   "frmMutasiChip.frx":834C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   45
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5310
      Left            =   0
      ScaleHeight     =   5310
      ScaleWidth      =   10860
      TabIndex        =   0
      Top             =   0
      Width           =   10860
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Select All"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   5940
         TabIndex        =   15
         Top             =   4995
         Width           =   4245
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Select All"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   4995
         Width           =   4245
      End
      Begin VB.CommandButton cmd 
         Height          =   480
         Index           =   3
         Left            =   5160
         Picture         =   "frmMutasiChip.frx":EB9E
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3210
         Width           =   555
      End
      Begin VB.CommandButton cmd 
         Height          =   480
         Index           =   1
         Left            =   5160
         Picture         =   "frmMutasiChip.frx":EC92
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2040
         Width           =   555
      End
      Begin VB.CommandButton cmdLink 
         Height          =   315
         Index           =   1
         Left            =   10320
         Picture         =   "frmMutasiChip.frx":ED84
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   315
         Width           =   405
      End
      Begin VB.CommandButton cmd 
         Height          =   480
         Index           =   2
         Left            =   5160
         Picture         =   "frmMutasiChip.frx":F10E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2640
         Width           =   555
      End
      Begin VB.CommandButton cmd 
         Height          =   480
         Index           =   0
         Left            =   5160
         Picture         =   "frmMutasiChip.frx":F200
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1455
         Width           =   555
      End
      Begin MSComctlLib.ListView List 
         Height          =   4395
         Index           =   0
         Left            =   105
         TabIndex        =   1
         Top             =   540
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   7752
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Prelot No"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Kd Brg"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nama Barang"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Qty"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Satuan"
            Object.Width           =   882
         EndProperty
      End
      Begin MSComctlLib.ListView List 
         Height          =   4230
         Index           =   1
         Left            =   5940
         TabIndex        =   4
         Top             =   705
         Width           =   4770
         _ExtentX        =   8414
         _ExtentY        =   7461
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Prelot No"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Kd Brg"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nama Barang"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Qty"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Satuan"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label lblWH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7650
         TabIndex        =   7
         Tag             =   "ASM"
         Top             =   300
         Width           =   2670
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "CHIP"
         Height          =   240
         Index           =   1
         Left            =   105
         TabIndex        =   8
         Top             =   300
         Width           =   4740
      End
      Begin VB.Line Line1 
         X1              =   8145
         X2              =   5940
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Gudang Tujuan"
         Height          =   285
         Index           =   0
         Left            =   5940
         TabIndex        =   5
         Top             =   360
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmMutasiChip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RsChip As New DBQuick
Private RsWH As New DBQuick
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
   
Private Sub PindahKanan(MoveAll As Boolean)
   Dim x, Y As Integer
   Y = List(1).ListItems.Count + 1
   For x = 1 To List(0).ListItems.Count
      If MoveAll Or List(0).ListItems(x).Checked Then
         List(1).ListItems.Add Y, List(0).ListItems(x).Text, List(0).ListItems(x).Text
         List(1).ListItems(Y).SubItems(1) = List(0).ListItems(x).SubItems(1)
         List(1).ListItems(Y).SubItems(2) = List(0).ListItems(x).SubItems(2)
         List(1).ListItems(Y).SubItems(3) = List(0).ListItems(x).SubItems(3)
         List(1).ListItems(Y).SubItems(4) = List(0).ListItems(x).SubItems(4)
         Y = Y + 1
      End If
   Next
   For x = List(0).ListItems.Count To 1 Step -1
      If MoveAll Or List(0).ListItems(x).Checked Then
         List(0).ListItems.Remove x
      End If
   Next
End Sub

Private Sub PindahKiri(MoveAll As Boolean)
   Dim x, Y As Integer
   Y = List(0).ListItems.Count + 1
   For x = 1 To List(1).ListItems.Count
      If MoveAll Or List(1).ListItems(x).Checked Then
         List(0).ListItems.Add Y, List(1).ListItems(x).Text, List(1).ListItems(x).Text
         List(0).ListItems(Y).SubItems(1) = List(1).ListItems(x).SubItems(1)
         List(0).ListItems(Y).SubItems(2) = List(1).ListItems(x).SubItems(2)
         List(0).ListItems(Y).SubItems(3) = List(1).ListItems(x).SubItems(3)
         List(0).ListItems(Y).SubItems(4) = List(1).ListItems(x).SubItems(4)
         Y = Y + 1
      End If
   Next
   
   For x = List(1).ListItems.Count To 1 Step -1
      If MoveAll Or List(1).ListItems(x).Checked Then
         List(1).ListItems.Remove x
      End If
   Next
End Sub

Private Sub chk_Click(Index As Integer)
   Dim x As Integer
   If Index = 0 Then
      If chk(Index).Value = 1 Then
         For x = 1 To List(0).ListItems.Count
            List(0).ListItems(x).Checked = True
         Next
      End If
   Else
      If chk(Index).Value = 1 Then
         For x = 1 To List(1).ListItems.Count
            List(1).ListItems(x).Checked = True
         Next
      End If
   End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0: PindahKanan True
        Case 1: PindahKanan False
        Case 2: PindahKiri False
        Case 3: PindahKiri True
    End Select
End Sub

Private Sub cmdLink_Click(Index As Integer)
   RsWH.DBOpen "Select Warehouse,[warehouse name],locations from warehouse", CNN
   If RsWH.DBRecordset.Recordcount > 0 Then
      Set mCall.FormData = RsWH.DBRecordset
      mCall.FromTagActive = "Warehouse"
   End If
End Sub



Private Sub CmdTombol_Click(Index As Integer)
   Dim x As Integer
   If Index = 0 Then
      Unload Me
   Else
      If MessageBox("Apakan Anda Ingin Menyimpan Data ?", "Konfirmasi", msgYesNo, msgQuestion) = 1 Then
         If List(1).ListItems.Count > 0 Then
            If lblWH.Caption = "" Then
               MessageBox "Gudang Tujuan Belum didefinisikan", "Warning", msgOkOnly
            Else
               For x = 1 To List(1).ListItems.Count
                  SendDataToServer "insert into [inventory tabel] (NoIdx,NoItem,Qty_In,stockTmp,typeTrans,sl_no,lokasiGdg) values (newID(),'" & _
                                                         List(1).ListItems(x).SubItems(1) & "'," & FQty(List(1).ListItems(x).SubItems(3)) & _
                                                         "," & FQty(List(1).ListItems(x).SubItems(3)) & ",'IP','" & List(1).ListItems(x).Text & "','" & lblWH.Caption & "')"
                  
                  SendDataToServer "update [mixing_header] set transfered = 1 where prelot = '" & List(1).ListItems(x).Text & "'"
               Next
               List(1).ListItems.Clear
            End If
         Else
            MessageBox "Tidak Data dalam daftar"
         End If
      End If
   End If
End Sub

Private Sub Form_Load()
   HiasFormManTell Picture1, Me
   LoadCHIP
   Set mCall = New frmCaller
End Sub

Private Sub LoadCHIP()
On Error GoTo ChipErr
   Dim x As Integer
   RsChip.DBOpen "select mixing_header.prelot,inventory.noItem, inventory.ItemName,mixing_header.total_sesudah,inventory.UOM from mixing_header inner join inventory on mixing_header.noItem = inventory.noItem where mixing_header.transfered=0", CNN
   List(0).ListItems.Clear
   With RsChip.DBRecordset
      If .Recordcount > 0 Then
         x = 1
         While Not .EOF
            List(0).ListItems.Add x, .Fields(0), .Fields(0)
            List(0).ListItems(x).SubItems(1) = .Fields(1)
            List(0).ListItems(x).SubItems(2) = .Fields(2)
            List(0).ListItems(x).SubItems(3) = .Fields(3)
            List(0).ListItems(x).SubItems(4) = .Fields(4)
            .MoveNext
            x = x + 1
         Wend
      End If
   End With
Exit Sub

ChipErr:
    MessageBox Err.Description, "Pengiriman Chip - LoadChip", msgOkOnly, msgExclamation
End Sub


Private Sub List_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
   If Item.Checked = False Then chk(Index).Value = 0
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
   lblWH.Caption = mCall.GetFieldByName(0)
End Sub
