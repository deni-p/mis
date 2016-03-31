VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FormFormulaEkstraksi1 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Konfigurasi Form Produksi"
   ClientHeight    =   4290
   ClientLeft      =   1980
   ClientTop       =   4305
   ClientWidth     =   10440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormFormulaEkstraksi1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   10440
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
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3615
      ScaleWidth      =   10440
      TabIndex        =   6
      Top             =   0
      Width           =   10440
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3760
         Picture         =   "FormFormulaEkstraksi1.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   1
         Tag             =   "Formula"
         Top             =   638
         Width           =   345
      End
      Begin VB.TextBox lblProses 
         Appearance      =   0  'Flat
         DataField       =   "Prosedur"
         DataSource      =   "MyDDE"
         Height          =   330
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "Formula"
         Top             =   630
         Width           =   2325
      End
      Begin MSDataGridLib.DataGrid GridKonfigurasi 
         Bindings        =   "FormFormulaEkstraksi1.frx":6BDC
         Height          =   3240
         Left            =   4320
         TabIndex        =   9
         Tag             =   "KP"
         Top             =   120
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   5715
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "ProsesID"
            Caption         =   "ID Proses"
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
            DataField       =   "Prosedur"
            Caption         =   "Prosedur"
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
               DividerStyle    =   6
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   6
            EndProperty
         EndProperty
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   1440
         X2              =   120
         Y1              =   930
         Y2              =   930
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1440
         X2              =   120
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label lblID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "EksNo"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1440
         TabIndex        =   3
         Tag             =   "Formula"
         Top             =   240
         Width           =   1905
      End
      Begin VB.Label lblNoEkstraksi 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAAF6F&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Ekstraksi"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   292
         Width           =   1035
      End
      Begin VB.Label lblHeaderForm 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAAF6F&
         BackStyle       =   0  'Transparent
         Caption         =   "Header Form"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   682
         Width           =   1050
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   3720
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   1005
      BindFormTAG     =   "Formula"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      BackColor       =   &H00EAAF6F&
      BackStyle       =   0  'Transparent
      Caption         =   "Header"
      Height          =   210
      Left            =   11040
      TabIndex        =   5
      Top             =   3840
      Width           =   585
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
      Left            =   4515
      TabIndex        =   4
      Top             =   285
      Width           =   495
   End
End
Attribute VB_Name = "FormFormulaEkstraksi1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private RcProses As New DBQuick

Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1

Private RcDetail As New DBQuick

Private RsDetail As DBQuick

Private MEdit As Boolean

Private rsBindControl As DBQuick

Private Sub bindControl(ByVal ParameterString As String)
On Error GoTo 1
  Set rsBindControl = New DBQuick
  If ParameterString = "" Then ParameterString = "11111111111" ': Exit Sub
  rsBindControl.DBOpen "SELECT ProdSample_Trans.prosedur From  ProdSample_Trans Where (ProdSample_Trans.typeTrans = '" & _
          ParameterString & "')", CNN, lckLockBatch
  ' Set MyDDE.ChildRecordset = rsBindControl.DBRecordset.Clone(adLockBatchOptimistic)
  'lblProses.Text = IIf(IsNull(MyDDE.ChildRecordset.Fields("prosedur")), "", MyDDE.ChildRecordset.Fields("prosedur"))
  lblProses.Text = IIf(IsNull(rsBindControl.Fields("prosedur")), "", rsBindControl.Fields("prosedur"))
  rsBindControl.CloseDB
Exit Sub
1:
MessageBox Err.Description, "frmformulaekstraksi1:bindcontrol" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub GridLayout()

  With GridKonfigurasi
    .Columns(0).width = 1200
    .Columns(1).width = 3600
  End With

End Sub

Private Function IndexAuto() As String
  On Error GoTo 2
  Dim Rc As New DBQuick
  Dim TglSaiki As String
  Dim Inom As String
  TglSaiki = Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2)
  Rc.DBOpen "SELECT  MAX(EksNo) AS MaxNom FROM         [ProdFormulaEkstraksi]", CNN, lckLockReadOnly

  With Rc

    If .DBRecordset.Recordcount <> 0 Then
      Inom = IIf(Not IsNull(.Fields(0)), Mid(.DBRecordset.Fields("MaxNom"), 5, 5), "0") + 1

      If Err.Number = 94 Then Inom = 1
    Else
      Inom = 1
    End If

    Select Case Len(Trim(Str(Inom)))

      Case 0
        IndexAuto = "KPR-" & Trim(Str(Inom))

      Case 1
        IndexAuto = "KPR-" & "0000" & Trim(Str(Inom))

      Case 2
        IndexAuto = "KPR-" & "000" & Trim(Str(Inom))

      Case 3
        IndexAuto = "KPR-" & "00" & Trim(Str(Inom))

      Case 4
        IndexAuto = "KPR-" & "0" & Trim(Str(Inom))
    End Select

  End With
Exit Function
2:
MessageBox Err.Description, "frmformulaekstraksi1:indexauto" & Err.Number, msgOkOnly, msgExclamation
End Function

'Private Function IndexAuto() As String
'   Dim Rc As New DBQuick
'   Dim TglSaiki As String
'   Dim Inom As String
'   TglSaiki = Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2)
'   Rc.DBOpen "SELECT  MAX(EksNo) AS MaxNom FROM         [ProdFormulaEkstraksi]", CNN, lckLockReadOnly
'
'   With Rc
'
'      If .DBRecordset.Recordcount <> 0 Then
'         Inom = IIf(Not IsNull(.Fields(0)), .Fields(0), "0") + 1
'      Else
'         Inom = 1
'      End If
'
'      Select Case Len(Trim(Str(Inom)))
'
'         Case 0
'            IndexAuto = "KPR-" & "0000" + Trim(Str(Inom))
'
'         Case 1
'            IndexAuto = "KPR-" & "0000" + Trim(Str(Inom))
'
'         Case 2
'            IndexAuto = "KPR-" & "0000" + Trim(Str(Inom))
'
'         Case 3
'            IndexAuto = "KPR-" & "0000" + Trim(Str(Inom))
'
'         Case 4
'            IndexAuto = "KPR-" & "0000" + Trim(Str(Inom))
'      End Select
'
'   End With
'
'End Function

Private Sub loadDetail()
On Error GoTo 3
  With MyDDE.ActiveRecordset

    If .Recordcount <> 0 Then
      Set RsDetail = New DBQuick

      If .EOF Or .BOF Then Exit Sub
      strSQL = _
              "SELECT ProdFormulaEkstraksi.EksNo, ProdFormulaEkstraksi.typeTrans, ProdProses.Prosedur,ProdFormulaEkstraksi_Detail.ProsesID From ProdFormulaEkstraksi_Detail  INNER JOIN ProdProses ON (ProdFormulaEkstraksi_Detail.ProsesID = ProdProses.ProsesID)  INNER JOIN ProdFormulaEkstraksi ON (ProdFormulaEkstraksi_Detail.EksNo = ProdFormulaEkstraksi.EksNo) Where  (ProdFormulaEkstraksi.EksNo ='" _
              & .Fields("eksno") & " ') "
      RsDetail.DBOpen strSQL, CNN
      Set MyDDE.ChildRecordset = RsDetail.DBRecordset.Clone(adLockBatchOptimistic)
      Set GridKonfigurasi.DataSource = MyDDE.ChildRecordset
    End If

  End With
Exit Sub
3:
MessageBox Err.Description, "frmformulaekstraksi1:loaddetail" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub OpenDetail(ByVal ParameterString As String)
On Error GoTo 4
  Set RcDetail = New DBQuick

  If ParameterString = "" Then ParameterString = "11111111111" ': Exit Sub
  RcDetail.DBOpen _
          "SELECT Prodconfigproses.ID, Prodconfigproses.ProsesID,Prodconfigproses.Analysis, Prodconfigproses.Methods, Prodconfigproses.MinValue, Prodconfigproses.MaxValue, Prodconfigproses.TypeOperator,  Prodconfigproses.UserName, Prodconfigproses.LastUpdate From  ProdAnalysis  INNER JOIN Prodconfigproses ON (ProdAnalysis.Analysis = Prodconfigproses.Analysis)  INNER JOIN ProdProses ON (Prodconfigproses.ProsesID = ProdProses.ProsesID) Where  (Prodconfigproses.ProsesID = " _
          & ParameterString & _
          ") GROUP BY  Prodconfigproses.ID,  Prodconfigproses.ProsesID,  Prodconfigproses.Analysis,  Prodconfigproses.Methods, Prodconfigproses.MinValue, Prodconfigproses.MaxValue, Prodconfigproses.TypeOperator, Prodconfigproses.UserName,Prodconfigproses.LastUpdate ", _
          CNN, lckLockBatch
  Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
  Set GridPenerimaanRL.DataSource = MyDDE.ChildRecordset
  RcDetail.CloseDB
Exit Sub
4:
MessageBox Err.Description, "frmformulaekstraksi1:opendetail" & Err.Number, msgOkOnly, msgExclamation

End Sub

Private Function OpenPartner(ByVal Index As Integer) As Boolean
On Error GoTo 5
  Set mCall = New frmCaller
    
  Select Case Index

      'Case 0: RcProses.DBOpen "SELECT Prodconfigproses.ID, Prodconfigproses.ProsesID,Prodconfigproses.Analysis, Prodconfigproses.Methods, Prodconfigproses.MinValue, Prodconfigproses.MaxValue,  Prodconfigproses.TypeOperator,  Prodconfigproses.UserName,  Prodconfigproses.LastUpdate From  ProdAnalysis  INNER JOIN Prodconfigproses ON (ProdAnalysis.Analysis = Prodconfigproses.Analysis)  INNER JOIN ProdProses ON (Prodconfigproses.ProsesID = ProdProses.ProsesID)", CNN, lckLockReadOnly
    Case 0
      RcProses.DBOpen "SELECT DISTINCT ProdSample_Trans.prosedur, ProdSample_Trans.TypeTrans From ProdSample_Trans ", CNN, _
              lckLockReadOnly

      ' Case 1: RcProses.DBOpen "SELECT ProdAnalysis.ID,ProdAnalysis.Analysis,ProdAnalysis.unit, ProdAnalysis.Accreditation From  ProdAnalysis  LEFT OUTER JOIN Prodconfigproses ON (ProdAnalysis.Analysis = Prodconfigproses.Analysis)  LEFT OUTER JOIN ProdProses ON (Prodconfigproses.ProsesID = ProdProses.ProsesID)", CNN, lckLockReadOnly
    Case 1
      RcProses.DBOpen _
              "SELECT DISTINCT Prodconfigproses.ProsesID,ProdProses.Prosedur,Prodconfigproses.Analysis,Prodconfigproses.Methods,Prodconfigproses.MinValue,Prodconfigproses.MaxValue From ProdProses  INNER JOIN Prodconfigproses ON (ProdProses.ProsesID = Prodconfigproses.ProsesID)  INNER JOIN ProdAnalysis ON (Prodconfigproses.Analysis = ProdAnalysis.Analysis) ORDER BY  Prodconfigproses.ProsesID ", _
              CNN, lckLockReadOnly

    Case 2
      RcProses.DBOpen "SELECT DISTINCT  ProdProses.Prosedur, ProdProses.ProsesID, ProdProses.keterangan From  ProdProses ", CNN, _
              lckLockReadOnly
  End Select
    
  If RcProses.Recordcount <> 0 Then

    Select Case Index

      Case 0
        mCall.FromTagActive = "HEADER FORM"

      Case 1
        mCall.FromTagActive = "KONFIGURASI PROSES"

      Case 2
        mCall.FromTagActive = "DAFTAR PROSEDUR"
    End Select

    Set mCall.FormData = RcProses.DBRecordset
    mCall.LookUp Me
  Else
    MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly
    OpenPartner = True
  End If
Exit Function
5:
MessageBox Err.Description, "frmformulaekstraksi1:openpartner" & Err.Number, msgOkOnly, msgExclamation
End Function

Private Sub PrepareSQL()
On Error GoTo 6
  With MyDDE
    .PrepareAppend = "insert into [ProdFormulaEkstraksi](EksNo,typeTrans) values ('" & lblid.Caption & "','" & lblHeader.Caption _
            & "')"
    .PrepareUpdate = "update [ProdFormulaEkstraksi] set EksNo='" & lblid.Caption & "', typeTrans= '" & lblHeader.Caption & _
            "' where EksNo='" & lblid.Caption & "'"
    .PrepareDelete = "delete from [ProdFormulaEkstraksi] where EksNo='" & lblid.Caption & "'" '.GetFieldByName("EksNo")
  End With
Exit Sub
6:
MessageBox Err.Description, "frmformulaekstraksi1:preparesql" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub SimpanDetail()
On Error GoTo 7
  With MyDDE.ChildRecordset

    If .Recordcount <> 0 Then
      .MoveFirst

      If SendDataToServer("DELETE FROM [ProdFormulaEkstraksi_Detail] WHERE     (EksNo = N'" & lblid.Caption & "')") = True Then

        Do

          If .EOF = True Then Exit Do
          SendDataToServer " INSERT INTO [ProdFormulaEkstraksi_Detail] (EksNo, ProsesID, typeTrans) " & " VALUES (N'" & _
                  lblid.Caption & "', N'" & .Fields("ProsesID") & "', N'" & lblHeader.Caption & "')"
          .MoveNext
        Loop

      End If
    End If

    .MoveLast
    GridKonfigurasi.Refresh
  End With
Exit Sub
7:
MessageBox Err.Description, "frmformulaekstraksi1:simpandetail" & Err.Number, msgOkOnly, msgExclamation

End Sub

Private Sub cmdLink_Click()
  OpenPartner 0
  'LoadDetail
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
  On Error GoTo 1
  ScanKey KeyCode, Shift, MyDDE
  If KeyCode = 27 Then Unload Me
Exit Sub
1:
MessageBox Err.Description, "Formformulaekstraksi:form_keydown" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Form_Load()
On Error GoTo 2
  With MyDDE
    .EditModeReplace = False
    Set .BindForm = Me
    .BindFormTAG = "KP"
    .SetPermissions = UserDeleteDenied
    Set .ActiveConnection = CNN
         
    ' .PrepareQuery = "select * from labProses "
         
    .PrepareQuery = "SELECT DISTINCT ProdFormulaEkstraksi.EksNo, ProdFormulaEkstraksi.typeTrans FROM ProdFormulaEkstraksi"
    ' Set GridKonfigurasi.DataSource = MyDDE.ActiveRecordset
  End With

  Set mCall = New frmCaller
  loadDetail
  
  HiasFormManTell Picture2, Me
Exit Sub
2:
MessageBox Err.Description, "formformulaekstraksi1:form_load" & Err.Number, msgOkOnly, msgExclamation
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

      Case 11

        GridKonfigurasi.MarqueeStyle = dbgFloatingEditor
        .AllowUpdate = True

      Case Else

        GridKonfigurasi.MarqueeStyle = dbgFloatingEditor
        .AllowUpdate = False
    End Select

  End With
Exit Sub
1:
MessageBox Err.Description, "formformulaekstraksi1:gridkonfigurasi_rowcolchange" & Err.Number, msgOkOnly, msgExclamation
End Sub





Private Sub mCall_RowColChange(ByVal TagForm As String, _
                               ByVal pRecordset As ADODB.Recordset)
On Error GoTo 1
  Select Case TagForm

    Case "HEADER FORM"

      With MyDDE
        lblProses.Text = mCall.GetFieldByName(0)
        lblHeader.Caption = mCall.GetFieldByName(1)
      End With

    Case ""

      With MyDDE
        .GetFieldByName("Analysis") = mCall.GetFieldByName(1)
        ' GridPenerimaanRL.Columns(1).Text = mCall.GetFieldByName(1)
      End With
              
    Case "KONFIGURASI PROSES"

      With MyDDE
        .ChildRecordset.MoveFirst
        srcText = mCall.GetFieldByName(0)

        Do

          If (Not .ChildRecordset.EOF) Or (Not .ChildRecordset.BOF) Or (mCall.GetFieldByName(0) = srcText) Then
            .ChildRecordset.Fields("Analysis") = mCall.GetFieldByName(2)
            .ChildRecordset.Fields("Methods") = mCall.GetFieldByName(3)
            .ChildRecordset.Fields("Minvalue") = mCall.GetFieldByName(4)
            .ChildRecordset.Fields("Maxvalue") = mCall.GetFieldByName(5)
            .ChildRecordset.MoveNext
          End If

        Loop

      End With
              
    Case "DAFTAR PROSEDUR"

      With MyDDE
        '.GetFieldByName("ProsesID") = mCall.GetFieldByName(1)
        'Debug.Print .ChildRecordset.Fields(0).Name
        .ChildRecordset.Fields("ProsesID") = mCall.GetFieldByName("ProsesID")
        .ChildRecordset.Fields("prosedur") = mCall.GetFieldByName("Prosedur")
      End With

  End Select
Exit Sub
1:
MessageBox Err.Description, "formformulaekstraksi:mcall_rowcolchage" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 1
  Select Case AdReasonActiveDb

    Case tmbDetail

      If mFirstCaller = False Then
        ' MyDDE.ChildRecordset.Fields("prosesID") = MyDDE.GetFieldByName("ProsesID")
        OpenPartner 2
        MEdit = True
      End If

    Case tmbAddNew

      MEdit = True
      lblid.Caption = IndexAuto
      lblProses.Enabled = False

    Case tmbSave

      If MyDDE.IsChildMemberReady = True Then
        SimpanDetail
        MEdit = False
      End If

    Case tmbCancel

      If MyDDE.ChildRecordset.Recordcount = 0 Then
        MEdit = False
        mVarDetailPOClose = False
      Else
        'DGPurchase.Columns(6).Visible = False
        'DGPurchase.Columns(7).Visible = True
      End If

    Case tmbEdit

      ' mEdit = True
  End Select
Exit Sub
1:
MessageBox Err.Description, "formformulaekstraksi1:mydde_afterprepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
                               ByVal pError As ADODB.Error, _
                               adStatus As ADODB.EventStatusEnum, _
                               ByVal pRecordset As ADODB.Recordset)
On Error GoTo 2:
  loadDetail
  On Error Resume Next

  With MyDDE
    lblid.Caption = IIf(IsNull(.GetFieldByName("EksNo")), "", .GetFieldByName("EksNo"))

    If .ChildRecordset.BOF Or .ChildRecordset.EOF Then Exit Sub
    lblProses.Text = IIf(IsNull(.ActiveRecordset.Fields("TypeTrans")), "", .ActiveRecordset.Fields("TypeTrans"))
    bindControl lblProses.Text
  End With
Exit Sub
2:
MessageBox Err.Description, "formformulaekstraksi1:mydde_movecomplete" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 3
  Select Case AdReasonActiveDb

    Case adAddNew
       
    Case tmbSave

      If MyDDE.CheckEmptyControl = False Then
        If CekGridKosong = False And MyDDE.ChildRecordset.Recordcount <> 0 Then
          MyDDE.IsChildMemberReady = True
          PrepareSQL
        Else
          MyDDE.IsChildMemberReady = False
          MyDDE.CancelTrans = True
        End If

      Else
        MyDDE.IsChildMemberReady = False
      End If
         
    Case 1

      cmdLink.Enabled = True
      cmdLink.SetFocus

    Case tmbDetail

      If MyDDE.CheckEmptyControl = False Then
      Else
        MyDDE.CancelTrans = mFirstCaller
      End If

    Case tmbCancel

      'MyDDE.CancelTrans = True
      MEdit = False
  End Select
Exit Sub
3:
MessageBox Err.Description, "formformulaekstraksi1:mydde_prepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

