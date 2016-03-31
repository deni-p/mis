VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "History"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11445
   Icon            =   "frmHistorySO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   11445
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   11415
      TabIndex        =   7
      Top             =   6840
      Width           =   11445
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   400
         Left            =   10125
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdDetail 
         Caption         =   "Detail"
         Height          =   375
         Left            =   45
         TabIndex        =   2
         Top             =   135
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   0
      ScaleHeight     =   6855
      ScaleWidth      =   11445
      TabIndex        =   4
      Top             =   0
      Width           =   11445
      Begin MSComctlLib.ListView ListHeader 
         Height          =   2535
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No SO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tanggal"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Customer"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Alamat"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cust PO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Batas Bayar"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Loco"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Currency"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Kurs"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Discount"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView listDetail 
         Height          =   3375
         Left            =   120
         TabIndex        =   1
         Top             =   3240
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   5953
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No Barang"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nama Barang"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Unit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Harga"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "PPN"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Total"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Tgl Kirim"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Daftar Sales Contract"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Data Detail"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   3000
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dsSO  As New DBQuick
Dim dsDetailSO As New DBQuick
Dim cHeaderCount As Integer
Dim cDetailCount As Integer
Private xID As String
Private xTgl As String
Private xPartner As String
Private aView As String
Private aTagName As String
Private aPartnerName As String
Private xForm As Form

Public Property Let PartnerName(vData As String)
   aPartnerName = vData
End Property


Public Property Let TagName(vData As String)
   aTagName = vData
End Property
Private Sub SetListViewColumn()
   Dim dsHeader As New DBQuick
   Dim ch As ColumnHeader
   ListHeader.ColumnHeaders.Clear
   strSQL = "Select column_name from fieldlist where table_name='" & aView & "'"
   dsHeader.DBOpen strSQL, CNN
   cHeaderCount = dsHeader.Recordcount
   If dsHeader.Recordcount > 0 Then
      dsHeader.MoveTopRecord
      For x = 1 To dsHeader.Recordcount
         If UCase(dsHeader.Fields(0)) = "NAMA" Then
            ListHeader.ColumnHeaders.Add x, "x" & x, aPartnerName
         Else
            ListHeader.ColumnHeaders.Add x, "x" & x, dsHeader.Fields(0)
         End If
         dsHeader.MoveNextRecord
      Next
   End If

   listDetail.ColumnHeaders.Clear
   strSQL = "Select column_name from fieldlist where table_name='" & aView & "Detail'"
   dsHeader.DBOpen strSQL, CNN
   cDetailCount = dsHeader.Recordcount
   If dsHeader.Recordcount > 0 Then
      dsHeader.MoveTopRecord
      For x = 1 To dsHeader.Recordcount - 1
         listDetail.ColumnHeaders.Add x, "x" & x, dsHeader.Fields(0)
         dsHeader.MoveNextRecord
      Next
   End If

End Sub

Public Property Let ViewName(vData As String)
   aView = vData
End Property


Public Property Let SONo(vData As String)
   xID = vData
End Property

Public Property Let RangeTanggal(vData As String)
   xTgl = vData
End Property


Public Property Let Partner(vData As String)
   xPartner = vData
End Property

Public Property Let DetailForm(vData As Form)
   Set xForm = vData
End Property

Private Sub cmdDetail_Click()
      If ListHeader.ListItems.Count > 0 Then
         xForm.IDParams = Trim(ListHeader.SelectedItem.Text)
         xForm.Show
      Else
         MessageBox "Data Tidak ada ! "
      End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Function GetStrSQL() As String
   strSQL = " SELECT * FROM " & aView
   If (Trim(xID) <> "") Or (Trim(xTgl) <> "") Or (Trim(xPartner) <> "") Then
      strSQL = strSQL & " where "
      If Trim(xID) <> "" Then
         strSQL = strSQL & " ID ='" & Trim(xID) & "'"
      End If
      If Trim(xTgl) <> "" Then
         If Trim(xID) <> "" Then strSQL = strSQL & " and "
         strSQL = strSQL & " Tanggal " & Trim(xTgl)
      End If
      If (Trim(xPartner)) <> "" Then
         If (Trim(xID) <> "") Or (Trim(xTgl) <> "") Then strSQL = strSQL & " and "
         strSQL = strSQL & " Company like '%" & xPartner & "%'"
      End If
   End If
   GetStrSQL = strSQL
End Function
Private Sub Form_Load()
   Dim lsv As ListItem
   'HiasForm Picture1, Me
   HiasFormManTell Picture2, Me
   Label1(0).Caption = "Daftar " & aTagName
   Caption = "History " & aTagName
   SetListViewColumn
   dsSO.DBOpen GetStrSQL, CNN, lckLockBatch, lckLockSync
   
   ListHeader.ListItems.Clear
   If dsSO.Recordcount > 0 Then
      dsSO.MoveTopRecord
      For x = 1 To dsSO.Recordcount
         Set lsv = ListHeader.ListItems.Add(x, "SO" & x, dsSO.Fields("ID"))
         For Y = 1 To cHeaderCount
            ListHeader.ListItems(x).ListSubItems.Add Y, "y" & Y, dsSO.Fields(Y)
         Next
        dsSO.MoveNextRecord
      Next
   End If
End Sub

Private Sub ListHeader_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Dim lsv As ListItem
   strSQL = "select * from " & aView & "Detail where ID='" & Trim(Item.Text) & "'"
   dsDetailSO.DBOpen strSQL, CNN
   
   listDetail.ListItems.Clear
   If dsDetailSO.Recordcount > 0 Then
      dsDetailSO.MoveTopRecord
      For x = 1 To dsDetailSO.Recordcount
         Set lsv = listDetail.ListItems.Add(x, "D" & x, dsDetailSO.Fields(0))
         For Y = 1 To cDetailCount
           listDetail.ListItems(x).ListSubItems.Add Y, "y" & Y, dsDetailSO.Fields(Y)
         Next
         dsDetailSO.MoveNextRecord
      Next
   End If
End Sub
