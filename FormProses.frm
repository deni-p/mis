VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FormProses 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Proses"
   ClientHeight    =   5190
   ClientLeft      =   3240
   ClientTop       =   3375
   ClientWidth     =   9405
   Icon            =   "FormProses.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   9405
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   0
      ScaleHeight     =   4575
      ScaleWidth      =   9405
      TabIndex        =   4
      Top             =   0
      Width           =   9405
      Begin MSDataGridLib.DataGrid GridProses 
         Bindings        =   "FormProses.frx":6852
         Height          =   3735
         Left            =   120
         TabIndex        =   3
         Tag             =   "KP"
         Top             =   825
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   6588
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Prosedur"
            Caption         =   "Prosedur"
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
            DataField       =   "Keterangan"
            Caption         =   "Keterangan"
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
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Pilih Jumlah Kolom"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   5535
         TabIndex        =   8
         Top             =   75
         Width           =   3705
         Begin VB.OptionButton OptYes 
            Appearance      =   0  'Flat
            BackColor       =   &H00EAAF6F&
            Caption         =   "1 (Satu) Kolom"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   255
            TabIndex        =   10
            Tag             =   "ana"
            Top             =   270
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1470
         End
         Begin VB.OptionButton OptNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00EAAF6F&
            Caption         =   "2 (Dua) Kolom"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   1770
            TabIndex        =   9
            Tag             =   "ana"
            Top             =   270
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "prosedur"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1155
         TabIndex        =   0
         Tag             =   "Proses"
         Top             =   90
         Width           =   2775
      End
      Begin VB.TextBox TxtKeterangan 
         Appearance      =   0  'Flat
         DataField       =   "keterangan"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1155
         TabIndex        =   1
         Tag             =   "Proses"
         Top             =   450
         Width           =   4290
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2610
         X2              =   90
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Label lblKeterangan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   945
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1635
         X2              =   90
         Y1              =   405
         Y2              =   405
      End
      Begin VB.Label lblProsedur 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   0
         Left            =   105
         TabIndex        =   6
         Top             =   105
         Width           =   720
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   4620
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   1005
      PrepareQuery    =   "select [ProsesID] as ID,[Prosedur],[keterangan] from [LabProses] order by ProsesID"
      BindFormTAG     =   "Proses"
      ActiveLanguage  =   1
   End
   Begin VB.Label lblID 
      Caption         =   "Label2"
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "FormProses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Nomor As Integer

Private Sub genID()
  On Error GoTo 3
  Dim rcSet As ADODB.Recordset
  strSQL = "select max(ProsesID) from [ProdProses]"

  If Me.Tag = "" Then
    Set rcSet = New ADODB.Recordset
    rcSet.Open strSQL, CNN, adOpenKeyset, adLockOptimistic

    If rcSet.Recordcount > 0 Then
      Nomor = "0" & rcSet.Fields(0)
      lblid.Caption = JadiAkhir(Nomor, 0)
    End If
  End If
Exit Sub
3:
MessageBox Err.Description, "frmproses:genid" & Err.Number, msgOkOnly, msgExclamation
End Sub

Function JadiAkhir(dariKode As Integer, jmlHEADER As Integer) As String
  JadiAkhir = Left(dariKode, jmlHEADER) + AkhirKode(Mid(dariKode, jmlHEADER + 1, 100))
End Function
Function AkhirKode(dariKode As String) As String
  On Error GoTo 1
  Dim A
  Dim I
  A = Val(dariKode) + 1

  For I = 1 To Len(dariKode)

    If Len(A) = I Then AkhirKode = Dgt(Len(dariKode) - I) & A
  Next I
Exit Function
1:
MessageBox Err.Description, "formproses:akhirkode" & Err.Number, msgOkOnly, msgExclamation
End Function

Function Dgt(jml As Currency) As String
On Error GoTo 2
  Dim I As Currency

  For I = 1 To jml
    Dgt = Dgt + "0"
  Next I
Exit Function
2:
MessageBox Err.Description, "formproses:dgt" & Err.Number, msgOkOnly, msgExclamation
End Function

Private Sub GridLayout()

  With GridProses
    .Columns(0).width = 4100  'ID
    .Columns(1).width = 3900 'prosedur
    '.Columns(2).Alignment = dbgCenter
    'Set .Columns(2).DataFormat = fmtGeneral
  End With

End Sub

Private Sub PrepareSQL()
On Error GoTo xErr
  With MyDDE
    .PrepareAppend = "insert into [ProdProses] (Prosedur,keterangan,kolom) values ('" & txtBox.Text & "','" & txtKeterangan.Text & "','" & IIf(OptYes(0).Value = True, "1", "0") & "')"
    .PrepareUpdate = "update [ProdProses] set kolom ='" & IIf(OptYes(0).Value = True, "1", "0") & "', Prosedur='" & txtBox.Text & "', keterangan= '" & txtKeterangan.Text & _
            "' where ProsesID=" & .GetFieldByName("ProsesID")
    .PrepareDelete = "delete from [ProdProses] where ProsesID=" & .GetFieldByName("ProsesID")
  End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub RefreshData()
On Error GoTo 4
  With MyDDE
    .EditModeReplace = False
    Set .BindForm = Me
    .BindFormTAG = "Proses"
    Set .ActiveConnection = CNN
    .PrepareQuery = "select [ProsesID],[Prosedur],[keterangan] from [ProdProses] order by ProsesID"

    If MyDDE.ActiveRecordset.Recordcount > 0 Then .ActiveRecordset.MoveLast
    Set GridProses.DataSource = .ActiveRecordset
  End With
Exit Sub
4:
MessageBox Err.Description, "formproses:refreshdata" & Err.Number, msgOkOnly, msgExclamation

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
On Error GoTo 1:
  ScanKey KeyCode, Shift, MyDDE

  If KeyCode = 27 Then Unload Me
Exit Sub
1:
MessageBox Err.Description, "formproses:form_keydown" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Form_Load()
  RefreshData
  'PrepareSQL
   
  HiasFormManTell Picture2, Me
  GridLayout
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 1
  Select Case AdReasonActiveDb

    Case tmbSave

      MyDDE.RefreshDatabase
      RefreshData

    Case tmbAddNew

      txtBox.SetFocus

    Case tmbEdit

      txtBox.Enabled = True

      'txtBox.SetFocus
    Case tmbPrint

      'CallRPTReport "Resources Type Table.rpt"
    Case 4

      RefreshData

    Case tmbDelete
      'mVarDataDc = False
      RefreshData
  End Select
Exit Sub
1:
MessageBox Err.Description, "formproses:mydde_afterprepareactivedb" & Err.Number, msgNo, msgExclamation
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
  PrepareSQL
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
                               ByVal pError As ADODB.Error, _
                               adStatus As ADODB.EventStatusEnum, _
                               ByVal pRecordset As ADODB.Recordset)
  ' txtBox.Text = IIf(IsNull(MyDDE.GetFieldByName("Prosedur")), "", MyDDE.GetFieldByName("Prosedur"))
  '  TxtKeterangan.Text = IIf(IsNull(MyDDE.GetFieldByName("keterangan")), "", MyDDE.GetFieldByName("keterangan"))
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 2

  Select Case AdReasonActiveDb

    Case tmbAddNew
txtBox.Text = ""
txtKeterangan.Text = ""
      genID
      PrepareSQL

    Case tmbSave

      genID
      PrepareSQL

      If txtKeterangan = "" Then
        txtKeterangan = " "

        If txtBox = "" Then
          MessageBox "Data belum lengkap.", "Peringatan", msgOkOnly
          MyDDE.IsChildMemberReady = False
          Exit Sub
        End If
      End If

      If MyDDE.CheckEmptyControl = False Then
        MyDDE.IsChildMemberReady = True
      Else
        MyDDE.IsChildMemberReady = False
      End If
      
    Case tmbDelete

      'INTERAL KONTROL BEFORE DELETE PROCESS ON
      'TRANSACTION & ListAnalysis FROM
      'If CekExist("LabSample_line", "Analysis", GridLab.Columns(1).Text) = False And _
       CekExist("LabListAnalysis", "Analysis", GridLab.Columns(1).Text) = False And _
       CekExist("LabSpecification", "Analysis", GridLab.Columns(1).Text) = False Then

      If MyDDE.CheckEmptyControl = False Then
        MyDDE.IsChildMemberReady = True
      Else
        MyDDE.IsChildMemberReady = False
      End If
         
  End Select
Exit Sub
2:
MessageBox Err.Description, "formproses:mydde_prepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

