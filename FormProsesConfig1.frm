VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FormProsesConfig1 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Konfigurasi Prosedur"
   ClientHeight    =   6960
   ClientLeft      =   2475
   ClientTop       =   2955
   ClientWidth     =   11610
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormProsesConfig1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   11610
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
      Height          =   6375
      Left            =   0
      ScaleHeight     =   6375
      ScaleWidth      =   11610
      TabIndex        =   3
      Tag             =   "pertama"
      Top             =   0
      Width           =   11610
      Begin MSDataGridLib.DataGrid GridKonfigurasi 
         Bindings        =   "FormProsesConfig1.frx":2A8B2
         Height          =   6135
         Left            =   2715
         TabIndex        =   8
         Tag             =   "KP"
         Top             =   165
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   10821
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   15
         RowDividerStyle =   6
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "Analysis"
            Caption         =   "Analisa"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Methods"
            Caption         =   "Methode"
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
            DataField       =   "MinValue"
            Caption         =   "Nilai Minimum"
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
         BeginProperty Column03 
            DataField       =   "MaxValue"
            Caption         =   "Nilai Maksimum"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "TypeOperator"
            Caption         =   "Operator"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "ID_ANALYSIS"
            Caption         =   "Id Analisa"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               DividerStyle    =   6
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   6
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               DividerStyle    =   6
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   14,74
            EndProperty
            BeginProperty Column05 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox lblProses 
         Appearance      =   0  'Flat
         DataSource      =   "MyDDE"
         Enabled         =   0   'False
         Height          =   345
         Left            =   11760
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "KP"
         Top             =   120
         Visible         =   0   'False
         Width           =   2325
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1860
         Top             =   2220
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
               Picture         =   "FormProsesConfig1.frx":2A8C6
               Key             =   "Orang"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FormProsesConfig1.frx":2B49A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FormProsesConfig1.frx":54AF4
               Key             =   "person1"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FormProsesConfig1.frx":553D0
               Key             =   "person2"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FormProsesConfig1.frx":55CAC
               Key             =   "TOP"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FormProsesConfig1.frx":56B00
               Key             =   "Dept"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView TVConfig 
         Height          =   5880
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   10372
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         SingleSel       =   -1  'True
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   0
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
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "PROSEDUR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   180
         Width           =   2505
      End
      Begin VB.Label lblPros 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Prosedur"
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
         Height          =   225
         Left            =   15240
         TabIndex        =   6
         Top             =   480
         Width           =   2505
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2610
         X2              =   120
         Y1              =   165
         Y2              =   165
      End
      Begin VB.Label lblProsedur 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAAF6F&
         BackStyle       =   0  'Transparent
         Caption         =   "Prosedur"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   5280
         TabIndex        =   4
         Top             =   660
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   6390
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   1005
      BindFormTAG     =   "KP"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   30
      Index           =   0
      Left            =   4515
      TabIndex        =   0
      Top             =   285
      Width           =   495
   End
End
Attribute VB_Name = "FormProsesConfig1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private mVarnode As Nodes

Private RcProses As New DBQuick

Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1

Private RcDetail As New DBQuick

Private rsDetail As DBQuick

Private MEdit As Boolean

Private Sub GridLayout()

  With GridKonfigurasi
    .Columns(0).width = 2200
    .Columns(1).width = 2800
    .Columns(2).width = 1000
    .Columns(3).width = 1000
    '.Columns(4).Width = 1000
    .Columns(5).width = 0
  End With

End Sub

Private Sub loadDetail()
On Error GoTo 1
  With MyDDE.ActiveRecordset

    If .Recordcount <> 0 Then
      Set rsDetail = New DBQuick

      If .EOF Or .BOF Then Exit Sub
      strSQL = _
              "SELECT DISTINCT Prodconfigproses.ID,Prodconfigproses.ProsesID, Prodconfigproses.ID_ANALYSIS, Prodconfigproses.Analysis,  Prodconfigproses.Methods,  Prodconfigproses.MinValue," _
              & _
              "Prodconfigproses.MaxValue, Prodconfigproses.TypeOperator, Prodconfigproses.UserName, Prodconfigproses.LastUpdate,ProdProses.Prosedur From  ProdProses  INNER JOIN Prodconfigproses ON (ProdProses.ProsesID = Prodconfigproses.ProsesID)  INNER JOIN ProdAnalysis ON (Prodconfigproses.Analysis = ProdAnalysis.Analysis) where Prodconfigproses.ProsesID='" _
              & .Fields("ProsesID") & "'"
      rsDetail.DBOpen strSQL, CNN
      Set MyDDE.ChildRecordset = rsDetail.DBRecordset.Clone(adLockBatchOptimistic)
      Set GridKonfigurasi.DataSource = MyDDE.ChildRecordset
    End If

  End With
Exit Sub
1:
MessageBox Err.Description, "formprosesconfig1:loaddetail" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub OpenDetail(ByVal ParameterString As String)
  Set RcDetail = New DBQuick
On Error GoTo 3
  If ParameterString = "" Then ParameterString = "11111111111" ': Exit Sub
  RcDetail.DBOpen _
          "SELECT Prodconfigproses.ID, Prodconfigproses.ProsesID,Prodconfigproses.Analysis, Prodconfigproses.Methods, Prodconfigproses.MinValue, Prodconfigproses.MaxValue, Prodconfigproses.TypeOperator,  Prodconfigproses.UserName, Prodconfigproses.LastUpdate From  ProdAnalysis  INNER JOIN Prodconfigproses ON (ProdAnalysis.Analysis = Prodconfigproses.Analysis)  INNER JOIN ProdProses ON (Prodconfigproses.ProsesID = ProdProses.ProsesID) Where  (Prodconfigproses.ProsesID = " _
          & ParameterString & _
          ") GROUP BY  Prodconfigproses.ID,  Prodconfigproses.ProsesID,  Prodconfigproses.Analysis,  Prodconfigproses.Methods, Prodconfigproses.MinValue, Prodconfigproses.MaxValue, Prodconfigproses.TypeOperator, Prodconfigproses.UserName,Prodconfigproses.LastUpdate ", _
          CNN, lckLockBatch
  Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
  Set GridKonfigurasi.DataSource = MyDDE.ChildRecordset
  RcDetail.CloseDB
Exit Sub
3:
MessageBox Err.Description, "formprosesconfig1:opendetail" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Function OpenPartner(ByVal Index As Integer) As Boolean
  Set mCall = New frmCaller
    
  Select Case Index

      'Case 0: RcProses.DBOpen "SELECT Prodconfigproses.ID, Prodconfigproses.ProsesID,Prodconfigproses.Analysis, Prodconfigproses.Methods, Prodconfigproses.MinValue, Prodconfigproses.MaxValue,  Prodconfigproses.TypeOperator,  Prodconfigproses.UserName,  Prodconfigproses.LastUpdate From  ProdAnalysis  INNER JOIN Prodconfigproses ON (ProdAnalysis.Analysis = Prodconfigproses.Analysis)  INNER JOIN ProdProses ON (Prodconfigproses.ProsesID = ProdProses.ProsesID)", CNN, lckLockReadOnly
    Case 0
      RcProses.DBOpen "SELECT DISTINCT ProdProses.ProsesID, ProdProses.Prosedur, ProdProses.keterangan From  ProdProses ", CNN, _
              lckLockReadOnly

    Case 1
      RcProses.DBOpen _
              "SELECT ProdAnalysis.Analysis,ProdAnalysis.unit,ProdAnalysis.ID From ProdAnalysis    order by  ProdAnalysis.Analysis", _
              CNN, lckLockReadOnly
      
  End Select
    
  If RcProses.Recordcount <> 0 Then

    Select Case Index

      Case 0
        mCall.FromTagActive = "MASTER PROSES"

      Case 1
        mCall.FromTagActive = "KONFIGURASI PROSES"
    End Select

    Set mCall.FormData = RcProses.DBRecordset
    mCall.LookUp Me
  Else
    MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly
    OpenPartner = True
  End If

End Function

Private Sub prepareSQL()
  
  With MyDDE
    .PrepareDelete = "delete from [Prodconfigproses] where ProsesID=" & lblProses.Tag
  End With

End Sub

Private Sub SimpanDetail()
On Error GoTo 5
  With MyDDE.ChildRecordset

    If .Recordcount <> 0 Then
      .MoveFirst

      If SendDataToServer("DELETE FROM [Prodconfigproses] WHERE  (ProsesID = '" & lblProses.Tag & "')") = True Then

        Do

          If .EOF = True Then Exit Do
          SendDataToServer _
                  " insert into [Prodconfigproses](ProsesID,ID_ANALYSIS,Analysis,Methods,MinValue,MaxValue,TypeOperator,UserName,LastUpdate) VALUES (" _
                  & lblProses.Tag & ", '" & .Fields("ID_ANALYSIS") & "', '" & .Fields("Analysis") & "', '" & .Fields("Methods") _
                  & "', '" & .Fields("MinValue") & "',' " & .Fields("MaxValue") & "', '" & .Fields("TypeOperator") & "',' " & _
                  "Neo" & "','" & Format(Now, "yyyy-MM-dd") & "')"
          .MoveNext
        Loop

      End If

      .MoveLast
      GridKonfigurasi.Refresh
    End If

  End With
Exit Sub
5:
MessageBox Err.Description, "formprosesconfig1:simpandetail" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
On Error GoTo 1
  ScanKey KeyCode, Shift, MyDDE

  If KeyCode = 27 Then Unload Me
Exit Sub
1:
MessageBox Err.Description, "formprosescinfig1:form_keydown" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Form_Load()
On Error GoTo 2
  HiasFormManTell Picture2, Me
  With MyDDE
    .EditModeReplace = False
    Set .BindForm = Me
    .BindFormTAG = "KP"
    .SetPermissions = UserDeleteDenied
    Set .ActiveConnection = CNN
    .PrepareQuery = _
            "SELECT DISTINCT Prodconfigproses.ID,Prodconfigproses.ProsesID,Prodconfigproses.ID_ANALYSIS,  Prodconfigproses.Analysis,  Prodconfigproses.Methods,  Prodconfigproses.MinValue," _
            & _
            "Prodconfigproses.MaxValue, Prodconfigproses.TypeOperator, Prodconfigproses.UserName, Prodconfigproses.LastUpdate,ProdProses.Prosedur From  ProdProses  INNER JOIN Prodconfigproses ON (ProdProses.ProsesID = Prodconfigproses.ProsesID)  INNER JOIN ProdAnalysis ON (Prodconfigproses.Analysis = ProdAnalysis.Analysis)"
    Set GridKonfigurasi.DataSource = MyDDE.ActiveRecordset
  End With

  Set mCall = New frmCaller
  Set mVarnode = TVConfig.Nodes
  TVConfig.Indentation = 300
  loadDetail
  GridLayout
  LoadTree
Exit Sub
2:
MessageBox Err.Description, "formprosesconfig1:form_load", msgOkOnly, msgExclamation
End Sub

Private Sub LoadTree()
  On Error GoTo 2
  Dim vNode As Node
  Dim rsForms As DBQuick
  Dim No  As Integer
  TVConfig.Nodes.Clear
  'Set vNode = TVConfig.Nodes.Add(, , "A", "All Forms")

  'vNode.Expanded = True
  'vNode.Tag = 0

  Set rsForms = New DBQuick
  rsForms.DBOpen "SELECT DISTINCT ProdProses.ProsesID, ProdProses.Prosedur, ProdProses.keterangan From  ProdProses order by ProdProses.Prosedur", CNN, _
          lckLockReadOnly
  No = 1

  If rsForms.Recordcount > 0 Then
    rsForms.DBRecordset.MoveFirst
    FirstNode = Trim(rsForms.DBRecordset.Fields(0))
    While Not rsForms.DBRecordset.EOF
      Set vNode = TVConfig.Nodes.Add(, , CStr(rsForms.DBRecordset.Fields(0)) & "A", Trim(IIf(IsNull(rsForms.DBRecordset.Fields( _
              "prosedur").Value), " ", rsForms.DBRecordset.Fields("prosedur").Value)), 2)
      vNode.Tag = No
      vNode.Expanded = True
      No = No + 1
      rsForms.DBRecordset.MoveNext
    Wend
  End If

  '  rsForms.DBRecordset.Close
  '  Set rsForms = Nothing
  '
  '  Set rsConfig = myPart.OpenDB("SELECT [REPORT_ID], [FIELD_NAME], [FIELD_TYPE] From report_filter WHERE ([REPORT_ID]= N'" & _
  '          UCase(FirstNode) & "')")
  '  Set GridConfig.DataSource = rsConfig
Exit Sub
2:
MessageBox Err.Description, "form prosesconfig1:loadtree" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub GridKonfigurasi_RowColChange(LastRow As Variant, _
                                         ByVal LastCol As Integer)
On Error GoTo 1
  If MEdit = False Then
    GridKonfigurasi.AllowUpdate = False
    GridKonfigurasi.MarqueeStyle = dbgFloatingEditor
    Exit Sub
  End If

  With GridKonfigurasi

    Select Case .col

      Case 1, 2, 3, 4, 5, 6

        GridKonfigurasi.MarqueeStyle = dbgFloatingEditor
        .AllowUpdate = True

      Case Else

        GridKonfigurasi.MarqueeStyle = dbgFloatingEditor
        .AllowUpdate = False
    End Select

  End With
Exit Sub
1:
MessageBox Err.Description, "formprosesconfig1:gridkonfigurasi_rowcolchange" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Label2_Click(Index As Integer)

End Sub

Private Sub lblPros_Click()

End Sub

Private Sub lblProsedur_Click(Index As Integer)

End Sub

Private Sub lblProses_Change()

End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, _
                               ByVal pRecordset As ADODB.Recordset)
On Error GoTo 1
  Select Case TagForm

    Case "MASTER PROSES"

      With MyDDE
        .GetFieldByName("ProsesID") = mCall.GetFieldByName(0)
        lblProses.Text = mCall.GetFieldByName(1)
        loadDetail
      End With

    Case ""

      With MyDDE
        .GetFieldByName("Analysis") = mCall.GetFieldByName("Analysis")
        ' GridKonfigurasi.Columns(1).Text = mCall.GetFieldByName(1)
      End With
              
    Case "KONFIGURASI PROSES"

      With MyDDE
        .ChildRecordset.Fields("Analysis") = mCall.GetFieldByName("Analysis")
        .ChildRecordset.Fields("ID_ANALYSIS") = mCall.GetFieldByName("ID")
      End With

  End Select
Exit Sub
1:
MessageBox Err.Description, "formprosesconfig1:mcall_rowcolchange" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 1
  Select Case AdReasonActiveDb

    Case tmbDetail

      If mFirstCaller = False Then
        MyDDE.ChildRecordset.Fields("prosesID") = MyDDE.GetFieldByName("ProsesID")

        If lblProses.Text = "" Then
          MessageBox "Pilih Prosedur dahulu", "Peringatan", msgOkOnly
          Exit Sub
        End If

        OpenPartner 1
        MEdit = True
      End If

    Case tmbAddNew

      MEdit = True
      lblProses.Enabled = False

    Case tmbSave

      If MyDDE.IsChildMemberReady = True Then
        SimpanDetail
        MEdit = False
      Else
        'MessageBox "Detail transaksi Purchase belum ada datanya.", "Peringatan", msgOkOnly
      End If

    Case tmbCancel

      If MyDDE.ChildRecordset.Recordcount = 0 Then
        MEdit = False
        mVarDetailPOClose = False
      Else
        'DGPurchase.Columns(6).Visible = False
        'DGPurchase.Columns(7).Visible = True
      End If

    Case tmbDelete
      SendDataToServer "DELETE FROM [Prodconfigproses] WHERE  (ProsesID = '" & lblProses.Tag & "')"

    Case tmbEdit

      ' mEdit = True
  End Select
Exit Sub
1:
MessageBox Err.Description, "formprosesconfig1:mydde_afterprepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
                               ByVal pError As ADODB.Error, _
                               adStatus As ADODB.EventStatusEnum, _
                               ByVal pRecordset As ADODB.Recordset)
  loadDetail

End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 2
  Select Case AdReasonActiveDb

    Case adAddNew
       
    Case tmbSave

      If lblProses.Text = "" Then
        MessageBox "Pilih Prosedur dahulu", "Peringatan", msgOkOnly
        MyDDE.IsChildMemberReady = False
        Exit Sub
      End If

      If MyDDE.CheckEmptyControl = False Then
        If CekGridKosong = False And MyDDE.ChildRecordset.Recordcount <> 0 Then
          MyDDE.IsChildMemberReady = True
          SimpanDetail
        Else
          MyDDE.IsChildMemberReady = False
          MyDDE.CancelTrans = True
        End If

      Else
        MyDDE.IsChildMemberReady = False
      End If
         
    Case 1

      If TVConfig.Nodes.Count > 0 Then
        lblProses.Tag = Mid(TVConfig.Nodes(1).Key, 1, Len(TVConfig.Nodes(1).Key) - 1)
      End If

      lblProses.Text = ""

    Case tmbDetail

      If MyDDE.CheckEmptyControl = False Then
      Else
        MyDDE.CancelTrans = mFirstCaller
      End If

    Case tmbDelete

      If MyDDE.CheckEmptyControl = False Then
        MyDDE.IsChildMemberReady = True
        prepareSQL
      Else
        MyDDE.IsChildMemberReady = False
      End If

    Case tmbCancel

      'MyDDE.CancelTrans = True
      MEdit = False
  End Select
Exit Sub
2:
MessageBox Err.Description, "formprosesconfig1:prepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Picture2_Click()

End Sub

Private Sub TVConfig_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo 1
  Select Case Node.Text

    Case Node.Key
      SearchDetail Mid(Node.Key, 1, Len(Node.Key) - 1)

    Case Else
      SearchDetail Mid(Node.Key, 1, Len(Node.Key) - 1)
      'TVConfig.SelectedItem.Bold = True
  End Select
Exit Sub
1:
MessageBox Err.Description, "formproseconfig1:tvconfig_nodeclick" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub SearchDetail(ByVal ParameterString As String)
On Error GoTo 4
  Set RcDetail = New DBQuick

  If ParameterString = "" Then ParameterString = "11111111111"
  strSQL = _
          "SELECT DISTINCT Prodconfigproses.ID,Prodconfigproses.ProsesID, Prodconfigproses.ID_ANALYSIS, Prodconfigproses.Analysis,  Prodconfigproses.Methods,  Prodconfigproses.MinValue," _
          & _
          "Prodconfigproses.MaxValue, Prodconfigproses.TypeOperator, Prodconfigproses.UserName, Prodconfigproses.LastUpdate,ProdProses.Prosedur From  ProdProses  INNER JOIN Prodconfigproses ON (ProdProses.ProsesID = Prodconfigproses.ProsesID)  INNER JOIN ProdAnalysis ON (Prodconfigproses.Analysis = ProdAnalysis.Analysis) where Prodconfigproses.ProsesID='" _
          & ParameterString & "'"
              
  RcDetail.DBOpen strSQL, CNN, lckLockBatch
  Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
  Set GridKonfigurasi.DataSource = MyDDE.ChildRecordset

  If MyDDE.ChildRecordset.Recordcount > 0 Then
    lblProses.Tag = Mid(TVConfig.SelectedItem.Key, 1, Len(TVConfig.SelectedItem.Key) - 1) '& TVConfig.SelectedItem.Text
    lblProses.Text = TVConfig.SelectedItem.Text
    ' SearchProsedur lblProses.Tag
  Else
    lblProses.Text = TVConfig.SelectedItem.Text
    lblProses.Tag = Mid(TVConfig.SelectedItem.Key, 1, Len(TVConfig.SelectedItem.Key) - 1)
  End If

  RcDetail.CloseDB
Exit Sub
4:
MessageBox Err.Description, "formprosesconfig1:seachdetail" & Err.Number, msgOkOnly, msgExclamation
End Sub

'Private Sub SearchProsedur(ByVal ParameterString As String)
'
'  Set RcDetail = New DBQuick
'
'  If ParameterString = "" Then ParameterString = "11111111111"
'  strSQL = "SELECT  ProdProses.Prosedur From ProdProses Where  ProdProses.ProsesID ='" & ParameterString & "'"
'  RcDetail.DBOpen strSQL, CNN, lckLockBatch
'  Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
'  lblProses = TVConfig.SelectedItem.Text 'TVConfig.SelectedItem.Text 'MyDDE.ChildRecordset.Fields("Prosedur")
'
'End Sub
