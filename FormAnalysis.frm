VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FormAnalysis 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form Master Analisa"
   ClientHeight    =   4800
   ClientLeft      =   3645
   ClientTop       =   3015
   ClientWidth     =   8970
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormAnalysis.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   8970
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
      Height          =   4095
      Left            =   0
      ScaleHeight     =   4095
      ScaleWidth      =   8970
      TabIndex        =   4
      Top             =   0
      Width           =   8970
      Begin VB.TextBox txtUOM 
         Appearance      =   0  'Flat
         DataField       =   "unit"
         DataSource      =   "MyDDE"
         Height          =   330
         Left            =   4500
         TabIndex        =   12
         Tag             =   "ana"
         Top             =   100
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Yes"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   7320
         TabIndex        =   1
         Tag             =   "ana"
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "No"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   8160
         TabIndex        =   2
         Tag             =   "ana"
         Top             =   120
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "analysis"
         DataSource      =   "MyDDE"
         Height          =   330
         Left            =   840
         TabIndex        =   0
         Tag             =   "ana"
         Top             =   100
         Width           =   2775
      End
      Begin MSDataGridLib.DataGrid GridLab 
         Bindings        =   "FormAnalysis.frx":6852
         Height          =   2940
         Left            =   120
         TabIndex        =   9
         Tag             =   "KP"
         Top             =   600
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   5186
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "ID"
            Caption         =   "ID"
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
            DataField       =   "Analysis"
            Caption         =   "Analisa"
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
            DataField       =   "unit"
            Caption         =   "Satuan"
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
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   6
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               DividerStyle    =   6
            EndProperty
         EndProperty
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   4785
         X2              =   3840
         Y1              =   415
         Y2              =   415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "**) Pilih No pada Grade Analisa untuk Analisa nilainya = Huruf"
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   3825
         Width           =   5370
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "*) Pilih Yes pada Grade Analisa untuk Analisa nilainya = Numeric/Angka"
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   3600
         Width           =   6135
      End
      Begin VB.Label lblid 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000005&
         Height          =   210
         Left            =   14640
         TabIndex        =   8
         Top             =   7560
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblAnalysis 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Analisa"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   150
         Width           =   540
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   3105
         X2              =   120
         Y1              =   405
         Y2              =   420
      End
      Begin VB.Label lblAccreditation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grade Analisa "
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   6120
         TabIndex        =   6
         Top             =   135
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label lblUOM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Satuan"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3840
         TabIndex        =   5
         Top             =   150
         Width           =   570
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   3
      Top             =   4230
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   1005
      PrepareQuery    =   "select * from LabAnalysis"
      BindFormTAG     =   "ana"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FormAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsLab As DBQuick
Dim yRow As Integer, xCol As Integer
Dim strSQL As String
Dim Nomor As Integer
Dim utility As New utility

Private MEdit As Boolean

Private Sub MoveCheckBox()
End Sub

Private Sub PrepareSQL()
   On Error GoTo xErr
    Dim isTrue As String

    With MyDDE
        isTrue = IIf(Option1(0).Value, "1", "0")
        .PrepareAppend = "insert into ProdAnalysis values ('" & txtBox.Text & "','" & txtUOM.Text & "', " & isTrue & ")"
        .PrepareUpdate = "update ProdAnalysis set Analysis='" & txtBox.Text & "',unit='" & txtUOM.Text & "',Accreditation=" & isTrue & " where ID=" & .GetFieldByName("ID")
        .PrepareDelete = "delete from ProdAnalysis where ID=" & .GetFieldByName("ID")
    End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub RefreshData()
    On Error GoTo FormAnalysis_RefreshData
    Set MyDDE.BindForm = Me
    Set MyDDE.ActiveConnection = CNN
    MyDDE.PrepareQuery = "select * from ProdAnalysis order by ID"

    If MyDDE.ActiveRecordset.Recordcount > 0 Then MyDDE.ActiveRecordset.MoveLast
    Set GridLab.DataSource = MyDDE.ActiveRecordset
Exit Sub
FormAnalysis_RefreshData:
    MessageBox Err.Description, "From Analysis : RefreshData" & Err.Number, msgOkOnly, msgExclamation
    
End Sub

Private Sub CheckCode_Click()
    On Error GoTo ClickErr
    Dim DataCheckBox As Integer

    If GridLab.Columns(GridLab.col).Value = True Then
        DataCheckBox = 1
    Else
        DataCheckBox = 0
    End If

    'DEFAULT "FALSE" VALUE FOR ACTIVE CHECKBOX
    'If CheckCode.Value = 0 Then
    '   If DataCheckBox <> CheckCode.Value Then
    '      GridLab.Columns(GridLab.Col).Value = False
    '   End If
    '   CheckCode.Caption = "NO"
    'Else
    ''DEFAULT "TRUE" VALUE FOR ACTIVE CHECKBOX
    '   If DataCheckBox <> CheckCode.Value Then
    '      GridLab.Columns(GridLab.Col).Value = True
    '   End If
    '   CheckCode.Caption = "YES"
    'End If
    Exit Sub

ClickErr:

    MsgBox Err.Description, vbCritical
End Sub

Private Sub CheckCode_KeyDown(KeyCode As Integer, _
                              Shift As Integer)
    On Error GoTo KeyErr

    Select Case KeyCode

        Case 13

            'GridLab.Columns(GridLab.Col).Value = CheckCode.Value
            'CheckCode.Visible = False
            'GridLab.SetFocus
        Case 27

            'CheckCode.Visible = False
            'GridLab.SetFocus
        Case Else
            'Exit Sub
   
    End Select

    Exit Sub

KeyErr:

    MsgBox Err.Description, vbCritical
End Sub

Private Sub CheckCode_LostFocus()
    'CheckCode.Visible = False
    'GridLab.SetFocus
End Sub

Private Sub Form_Activate()
'    Gelas False
 '   LockObject Me, True
  '  BarContent "Analysis Data", PanelDesc
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo formanalysis_form_keydown
    ScanKey KeyCode, Shift, MyDDE

    If KeyCode = 27 Then Unload Me
Exit Sub
formanalysis_form_keydown:
MessageBox Err.Description, "formanalysis:form_keydown" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Form_Load()
    RefreshData
    HiasFormManTell Picture2, Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    'CheckCode.Visible = False
    'GridLab.SetFocus

    'Utility.CheckPending rsLab
   ' BarContent "", ""
End Sub

Private Sub GridLab_Click()
On Error GoTo 1
If xCol = -1 And yRow = -1 Then utility.SelectGrid GridLab, xCol, yRow
Exit Sub
1:
    MessageBox Err.Description, "FromAnalysis:Gridlab_click" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub GridLab_HeadClick(ByVal ColIndex As Integer)
    ' MyDDE.ActiveRecordset.Sort = GridLab.Columns(ColIndex).Caption
End Sub

Private Sub GridLab_KeyDown(KeyCode As Integer, _
                            Shift As Integer)
    'Utility.ScanKey Me, rsLab, GridLab, KeyCode, Shift
End Sub

Private Sub GridLab_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              x As Single, _
                              Y As Single)
On Error GoTo 2
    xCol = GridLab.ColContaining(x)
    yRow = GridLab.RowContaining(Y)

    If xCol = -1 And yRow = -1 Then
        GridLab.ToolTipText = " Select all record "
    Else
        GridLab.ToolTipText = ""
    End If
Exit Sub
2:
MessageBox Err.Description, "FromAnalysis : Gridlab_Mousemove" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub GridLab_Scroll(Cancel As Integer)
    MoveCheckBox
End Sub


Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
    On Error GoTo 1

    Select Case AdReasonActiveDb

        Case tmbSave

            MyDDE.RefreshDatabase
            RefreshData

        Case tmbAddNew
            txtUOM.Text = ""
            txtBox.Text = ""
            'mVarDataDc = True
            txtBox.SetFocus

            If IsEmpty(MyDDE.GetFieldByName("Accreditation")) Then
                If Option1(0).Value = True Then
                    MyDDE.ActiveRecordset.Fields("Accreditation").Value = 1
                Else
                    MyDDE.ActiveRecordset.Fields("Accreditation").Value = 0
                End If
        
            End If

        Case tmbEdit

            MEdit = True

            If txtBox.Enabled = False Then
                txtBox.Enabled = True
                txtBox.SetFocus
            End If

        Case tmbDelete

            RefreshData

        Case tmbPrint

            CallRPTReport "Resources Type Table.rpt"

        Case 4

            RefreshData

        Case Else
            'mVarDataDc = False
    End Select
Exit Sub
1:
MessageBox Err.Description, "frmanalysis:mydde_afterprepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 2
    If AdReasonActiveDb = tmbSave Then PrepareSQL
    PrepareSQL
Exit Sub
2:
MessageBox Err.Description, "frmanalysis_mydde_executeorder" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
                               ByVal pError As ADODB.Error, _
                               adStatus As ADODB.EventStatusEnum, _
                               ByVal pRecordset As ADODB.Recordset)
On Error GoTo 3
    lblid.Caption = IIf(IsNull(MyDDE.GetFieldByName("ID")), "", MyDDE.GetFieldByName("ID"))

    If MyDDE.GetFieldByName("Accreditation") = True Then
        Option1(0).Value = True
    Else
        Option1(1).Value = True
        Exit Sub
    End If
  
    txtUOM.Text = IIf(IsNull(MyDDE.GetFieldByName("unit")) Or MyDDE.GetFieldByName("unit") = " ", "-", MyDDE.GetFieldByName("unit"))
Exit Sub
3:
MessageBox Err.Description, "frmanalysis:mydde_movecomplete" & Err.Number, msgOkOnly, msgExclamation
End Sub
   
Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 4
    Select Case AdReasonActiveDb

        Case tmbAddNew

            'genID
            'prepareSQL
            ' IIf Option1(0).Value = True, GridLab.Columns(2).Text = "Yes", GridLab.Columns(2).Text = "No"
       
        Case tmbSave

            If txtUOM = "" Then
                txtUOM = " "

                If txtBox = "" Then
                    MessageBox "Data belum lengkap.", "Peringatan", msgOkOnly
                    MyDDE.IsChildMemberReady = False
                    Exit Sub
                End If
            End If

            If MyDDE.CheckEmptyControl = False Then
                MyDDE.IsChildMemberReady = True
                PrepareSQL
            Else
                MyDDE.IsChildMemberReady = False
            
            End If

        Case tmbDelete

            If MyDDE.CheckEmptyControl = False Then
                MyDDE.IsChildMemberReady = True
            Else
                MyDDE.IsChildMemberReady = False
                MyDDE.CancelTrans = True
            End If

    End Select

Exit Sub
4:
MessageBox Err.Description, "formanalysis:mydde_prepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub TxtFind_Change()
    'With datalab.Recordset
    '   .Find "Analysis = '" & txtFind.Text & "'", , adSearchForward, 1
    '   If .EOF Then
    '       .MoveLast
    '   Else
    '      Utility.Block txtFind
    '   End If
    'End With
End Sub

