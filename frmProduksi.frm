VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProduksi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Proses Produksi"
   ClientHeight    =   4740
   ClientLeft      =   15
   ClientTop       =   315
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProduksi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7455
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   0
      ScaleHeight     =   4935
      ScaleWidth      =   7455
      TabIndex        =   16
      Top             =   4050
      Width           =   7455
      Begin VB.CommandButton CmdTombol 
         Caption         =   "E&xit"
         Height          =   555
         Index           =   0
         Left            =   6660
         Picture         =   "frmProduksi.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   0
      ScaleHeight     =   4050
      ScaleWidth      =   7455
      TabIndex        =   14
      Top             =   0
      Width           =   7455
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "City"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   0
         Left            =   1635
         MaxLength       =   50
         TabIndex        =   0
         Tag             =   "Partner"
         Top             =   75
         Width           =   2805
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "City"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   5
         Left            =   1635
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "Partner"
         Top             =   420
         Width           =   2805
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Dryer"
         Enabled         =   0   'False
         Height          =   435
         Index           =   12
         Left            =   5160
         TabIndex        =   11
         Top             =   2220
         Width           =   2010
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Crusher"
         Enabled         =   0   'False
         Height          =   435
         Index           =   13
         Left            =   5145
         TabIndex        =   12
         Top             =   2745
         Width           =   2010
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Penjemuran"
         Enabled         =   0   'False
         Height          =   435
         Index           =   11
         Left            =   5160
         TabIndex        =   10
         Top             =   1695
         Width           =   2010
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Cutter"
         Enabled         =   0   'False
         Height          =   435
         Index           =   10
         Left            =   5160
         TabIndex        =   9
         Top             =   1170
         Width           =   2010
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Hydraulic Press"
         Enabled         =   0   'False
         Height          =   435
         Index           =   9
         Left            =   2865
         TabIndex        =   8
         Top             =   3300
         Width           =   2010
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Concrete Press"
         Enabled         =   0   'False
         Height          =   435
         Index           =   8
         Left            =   2865
         TabIndex        =   7
         Top             =   2745
         Width           =   2010
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Pembungkusan"
         Enabled         =   0   'False
         Height          =   435
         Index           =   7
         Left            =   2865
         TabIndex        =   6
         Top             =   2220
         Width           =   2010
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Gellification"
         Enabled         =   0   'False
         Height          =   435
         Index           =   6
         Left            =   2865
         TabIndex        =   5
         Top             =   1680
         Width           =   2010
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Filter Press"
         Enabled         =   0   'False
         Height          =   435
         Index           =   5
         Left            =   2865
         TabIndex        =   4
         Top             =   1170
         Width           =   2010
      End
      Begin VB.CommandButton cmdLink 
         Height          =   315
         Index           =   0
         Left            =   4455
         Picture         =   "frmProduksi.frx":834C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   75
         Width           =   390
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1740
         Top             =   2820
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   16777215
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProduksi.frx":86D6
               Key             =   "Orang"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProduksi.frx":92AA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProduksi.frx":32904
               Key             =   "person1"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProduksi.frx":331E0
               Key             =   "person2"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProduksi.frx":33ABC
               Key             =   "TOP"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProduksi.frx":34910
               Key             =   "Dept"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView TVConfig 
         Height          =   2340
         Left            =   120
         TabIndex        =   3
         Top             =   1425
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   4128
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         SingleSel       =   -1  'True
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   0
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
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   7305
         X2              =   75
         Y1              =   930
         Y2              =   930
      End
      Begin VB.Label lblMetode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   4440
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Rekomendasi"
         Height          =   195
         Left            =   105
         TabIndex        =   19
         Top             =   120
         Width           =   1245
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   1845
         X2              =   75
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lblPROSEDURPRODUKSI 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "REKOMENDASI EKSTRAKSI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   1155
         Width           =   2505
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1740
         X2              =   75
         Y1              =   375
         Y2              =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Manufacture Order"
         Height          =   195
         Left            =   105
         TabIndex        =   15
         Top             =   510
         Width           =   1380
      End
   End
   Begin VB.Label lblNoRl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label lblSplNo 
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmProduksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1

Private RcDetail As DBQuick

Private RcProduksi As New DBQuick

Private Sub cmd_Click(Index As Integer)

    Select Case Index

        Case 1
            FrmProdAcidTreatment.SetFocus

        Case 2
            frmProdBleaching.SetFocus

        Case 3


        Case 4


        Case 5
            FrmProdFilterPress.SetFocus

        Case 6
            'FrmProdGellification.SetFocus
             FrmGELL.SetFocus
        Case 7
            frmprodpembungkusan.SetFocus

        Case 8

            frmPconcrete.SetMode = 0
            frmPconcrete.SetFocus

        Case 9

            frmpress.SetMode = 0
            frmpress.SetFocus

        Case 10

            FrmPcutter.SetMode = 0
            FrmPcutter.SetFocus

        Case 11

            FrmJemur.SetMode = 0
            FrmJemur.SetFocus

        Case 12

            frmDryer.SetMode = 0
            frmDryer.SetFocus

        Case 13

            frmPCrusher.SetMode = 0
            frmPCrusher.SetFocus
    End Select

End Sub

Private Sub cmdLink_Click(Index As Integer)
    OpenPartner 0
End Sub

Private Sub LoadTree(ByVal ParameterString As String)
    Dim vNode As Node
    Dim rsForms As DBQuick
    Dim No  As Integer
    tvConfig.Nodes.Clear

    Set rsForms = New DBQuick
    rsForms.DBOpen "SELECT FORMID,FormName From view_rekomekstraksi_proses where splNo='" & txtBox(0) & "'", CNN, lckLockBatch
    No = 1
    If rsForms.Recordcount > 0 Then
        rsForms.DBRecordset.MoveFirst
        FirstNode = Trim(rsForms.DBRecordset.Fields(0))
        While Not rsForms.DBRecordset.EOF
            Set vNode = tvConfig.Nodes.Add(, , CStr(rsForms.DBRecordset.Fields("FORMID")) & "A", Trim(IIf(IsNull(rsForms.DBRecordset.Fields("FormName").Value), " ", rsForms.DBRecordset.Fields("FormName").Value)), 2)
            vNode.Tag = No
            vNode.Expanded = True
            No = No + 1
            rsForms.DBRecordset.MoveNext
        Wend
    End If

End Sub

Private Sub LoadButton(ByVal ParameterString As String)
    Dim vNode As Node
    Dim rsForms As DBQuick
    Dim No  As Integer
    'TVConfig.Nodes.Clear
    Dim x As Integer
    Set rsForms = New DBQuick

    rsForms.DBOpen "SELECT [Manufacture Order].OrderID,[Manufacture Order].ekstraksi_no,[Manufacture Order].no_rekomendasi,[Order Output Detail].WCID, wcenter_header.formid, labformrekomendasi.FormName From [Order Output Detail] INNER JOIN [Manufacture Order] ON ([Order Output Detail].OrderID = [Manufacture Order].OrderID) INNER JOIN wcenter_header ON ([Order Output Detail].WCID = wcenter_header.WCID) " & " INNER JOIN labformrekomendasi ON (wcenter_header.formid = labformrekomendasi.formid) Where [Manufacture Order].OrderID = '" & ParameterString & "'", CNN, lckLockBatch
    No = 1
    Do While Not rsForms.DBRecordset.EOF

        For x = 5 To 13
        If UCase(rsForms.DBRecordset.Fields("FormName")) = UCase(cmd(x).Caption) Then cmd(x).Enabled = True
        Next
    rsForms.DBRecordset.MoveNext
    Loop
    DoEvents
    rsForms.CloseDB
End Sub

Private Sub EnableProses(ByVal ParameterString As String)
    Dim ncount As Integer
    Dim nRecord As Integer
    
    Set RcDetail = New DBQuick

    If ParameterString = "" Then ParameterString = "11111111111" ': Exit Sub
  
    RcDetail.DBOpen "SELECT LabRekomEkstraksi.SplNo,LabRekomEkstraksi_Line.FORMID, LabRekomEkstraksi_Line.FormName From LabRekomEkstraksi_Line INNER JOIN LabRekomEkstraksi ON (LabRekomEkstraksi_Line.SplNo = LabRekomEkstraksi.SplNo) Where  LabRekomEkstraksi.SplNo = '" & ParameterString & "'", CNN, lckLockBatch

    For nRecord = 0 To RcDetail.Recordcount
        For ncount = 0 To 12
            Debug.Print UCase(cmd(ncount).Caption)

            If UCase(cmd(ncount).Caption) = RcDetail.Fields("FORMNAME") Then
                cmd(ncount).Enabled = True
            End If

        Next ncount

        RcDetail.MoveNextRecord
    Next nRecord

    RcDetail.CloseDB
   
End Sub

Private Function OpenPartner(ByVal Index As Integer) As Boolean
    Set mCall = New frmCaller
    
    Select Case Index
        Case 0: RcProduksi.DBOpen "select [manufacture Order].orderID,[manufacture Order].no_rekomendasi, Case labrekomekstraksi.tempatalkali WHEN 0 THEN 'Reaktor' WHEN 1 THEN 'Bak Luar' WHEN 2 THEN 'AutoClave'  END As [Tempat Ekstraksi],labrekomekstraksi.rlno, labrekomekstraksi.methode  From [manufacture Order]  inner join labrekomekstraksi on [manufacture Order].no_rekomendasi =   labrekomekstraksi.splno where [manufacture Order].status = 'RELEASED'", CNN
        
    End Select
    
    If RcProduksi.Recordcount <> 0 Then

        Select Case Index

            Case 0
                mCall.FromTagActive = "PRODUKSI"

        End Select

        Set mCall.FormData = RcProduksi.DBRecordset
        mCall.LookUp Me
    Else

        MessageBox "Rekomendasi Produksi Masih Kosong", "Peringatan", msgOkOnly, msgCrtical
        OpenPartner = True
    End If

End Function

Private Sub mCall_RowColChange(ByVal TagForm As String, _
                               ByVal pRecordset As ADODB.Recordset)
    Dim rsMO As New DBQuick
    Dim x As Integer

    Select Case TagForm

        Case "PRODUKSI"
            txtBox(5) = mCall.GetFieldByName("OrderID")
            txtBox(0) = mCall.GetFieldByName("no_rekomendasi")
            lblNoRL.Caption = mCall.GetFieldByName("rlno")
            lblMetode.Caption = mCall.GetFieldByName("methode")
            lblSplNo.Caption = mCall.GetFieldByName("rlno")
            If (txtBox(5).Text <> "") And (txtBox(0).Text <> "") Then
                tvConfig.Enabled = True
                '*** seleksi semua tombol & treeview yang aktif ***
                LoadTree txtBox(0).Text
                LoadButton txtBox(5).Text
                '***
                For x = 5 To 13
                    cmd(x).Enabled = True
                Next
            
            Else

                For x = 5 To 13
                    cmd(x).Enabled = False
                Next

                tvConfig.Enabled = False
            End If

    End Select

    rsMO.CloseDB
    
End Sub

Private Sub CmdTombol_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    Set mCall = New frmCaller
    HiasFormManTell Picture1, Me
   lblPROSEDURPRODUKSI(0).BackColor = &H8000000C
End Sub

Private Sub TVConfig_DblClick()

    If tvConfig.Nodes.Count < 1 Then Exit Sub

    Select Case tvConfig.SelectedItem.Text
    
        Case "ACID TREATMENT"
            frmProdAcid.SetFocus

        Case "ALKALI TREATMENT"
            frmProdAlkali.SetFocus

        Case "BLEACHING TREATMENT"
            frmProdBleaching.SetFocus

        Case "EXTRACTION REAKTOR"
            FrmProdEksReaktor.SetFocus

        Case "EXTRACTION AUTOCLAVE"
            FrmProdEkAutoClave.SetFocus
            
        Case "FILTER PRESS"
            FrmProdFilterPress.SetFocus
    End Select

End Sub


Private Sub txtBox_KeyPress(Index As Integer, KeyAscii As Integer)
   If (Index = 0) And (KeyAscii = 13) Then
      RcProduksi.DBOpen "select [manufacture Order].orderID,[manufacture Order].no_rekomendasi, Case labrekomekstraksi.tempatalkali WHEN 0 THEN 'Reaktor' WHEN 1 THEN 'Bak Luar' WHEN 2 THEN 'AutoClave'  END As [Tempat Ekstraksi],labrekomekstraksi.rlno, labrekomekstraksi.methode  From [manufacture Order]  inner join labrekomekstraksi on [manufacture Order].no_rekomendasi = labrekomekstraksi.splno where [manufacture Order].status = 'RELEASED'", CNN, lckLockBatch
      If RcProduksi.DBRecordset.Recordcount > 0 Then
         txtBox(5).Text = RcProduksi.DBRecordset.Fields("OrderID")
         txtBox(0).Text = RcProduksi.DBRecordset.Fields("no_rekomendasi")
         lblNoRL.Caption = RcProduksi.DBRecordset.Fields("rlno")
         lblMetode.Caption = RcProduksi.DBRecordset.Fields("methode")
         lblSplNo.Caption = RcProduksi.DBRecordset.Fields("rlno")
         If (txtBox(5).Text <> "") And (txtBox(0).Text <> "") Then
             tvConfig.Enabled = True
             '*** seleksi semua tombol & treeview yang aktif ***
             LoadTree txtBox(0).Text
             LoadButton txtBox(5).Text
             '***
             For x = 5 To 13
                 cmd(x).Enabled = True
             Next
         Else
             For x = 5 To 13
                 cmd(x).Enabled = False
             Next
             tvConfig.Enabled = False
         End If

      End If
   End If
End Sub
