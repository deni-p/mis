VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEntryForm 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parameter Filter Laporan"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEntryForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6255
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chk 
      BackColor       =   &H00EAAF6F&
      Caption         =   "Select All"
      Height          =   330
      Left            =   120
      TabIndex        =   4
      Top             =   3945
      Width           =   1335
   End
   Begin VB.PictureBox PicTombol 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   6255
      TabIndex        =   7
      Top             =   4305
      Width           =   6255
      Begin VB.CommandButton CmdTombol 
         Caption         =   "E&xit"
         Height          =   555
         Index           =   0
         Left            =   9030
         Picture         =   "frmEntryForm.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   100
         Width           =   720
      End
      Begin VB.Frame FrTombol 
         BackColor       =   &H00C0FFFF&
         Height          =   30
         Left            =   -45
         TabIndex        =   8
         Top             =   0
         Width           =   9945
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Save"
         Height          =   555
         Index           =   0
         Left            =   75
         Picture         =   "frmEntryForm.frx":834C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   90
         Width           =   720
      End
      Begin VB.CommandButton cmd 
         Caption         =   "E&xit"
         Height          =   555
         Index           =   1
         Left            =   795
         Picture         =   "frmEntryForm.frx":EB9E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   90
         Width           =   720
      End
   End
   Begin MSDataListLib.DataCombo cmbForm 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   90
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   714
      _Version        =   393216
      Style           =   2
      ListField       =   "Alias Report"
      BoundColumn     =   "NoIdx"
      Text            =   "cmbForm"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView LVField 
      Height          =   3165
      Left            =   105
      TabIndex        =   3
      Top             =   780
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5583
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Field Name"
         Object.Width           =   5009
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Data Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pilih field dibawah untuk menentukan filter Laporan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   4350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Laporan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data was Inserted"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4455
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   1560
   End
End
Attribute VB_Name = "frmEntryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mVarOpMode As String
Private mVarRPTName As String
Private myPart As New utility

Private mVarFormName As String
Private RcListReport As New Recordset
Private RcDetReport As New DBQuick
Dim rsForm As ADODB.Recordset
Dim strSQL As String
'Public Property Let ReportName(ByVal vData As String)
'    mVarRPTName = vData
'End Property
Public Property Let OperationMode(ByVal vData As String)
    mVarOpMode = vData
End Property

Public Property Let FormName(ByVal vData As String)
    mVarFormName = Trim(vData)
End Property

Private Function GetFieldType(ByVal dataType As String) As Integer

    '11  Bit
    '135 DateTime
    '202 nvarchar
    '203 ntext
    '3   int
    '5   Float
    '6   Money
    Select Case UCase(dataType)

        Case "BIT"

            GetFieldType = 11

        Case "DATETIME"

            GetFieldType = 135

        Case "NVARCHAR"

            GetFieldType = 202

        Case "NTEXT"

            GetFieldType = 203

        Case "INT"

            GetFieldType = 3

        Case "FLOAT"

            GetFieldType = 5

        Case "MONEY"

            GetFieldType = 6
    End Select

End Function

Private Sub chk_Click()
    Dim x As Integer

    If chk.Value = 1 Then

        For x = 1 To LVField.ListItems.Count
            LVField.ListItems.Item(x).Checked = True
        Next

    ElseIf chk.Value = 0 Then

        For x = 1 To LVField.ListItems.Count
            LVField.ListItems.Item(x).Checked = False
        Next

    End If

End Sub

Private Sub cmbForm_Change()
    Dim rsField As ADODB.Recordset, rsValidate As ADODB.Recordset
    Dim rsCheck As New ADODB.Recordset
    Dim strCheck As String
    Dim vLst As ListItem
    Dim x As Integer
    Dim Y As Integer
    LVField.ListItems.Clear
    On Error GoTo xErr
    If (mVarOpMode = "Insert") Then
        Set rsValidate = myPart.OpenDB("SELECT Report_Filter.FIELD_NAME, Report_Filter.FIELD_TYPE, Report_Filter.OBJECT_TYPE From  [Report Modules] INNER JOIN Report_Filter ON ([Report Modules].NoIdx = Report_Filter.REPORT_ID) Where  Report_Filter.REPORT_ID ='" & cmbForm.BoundText & "' ORDER BY Report_Filter.FIELD_NAME")
        
        If (rsValidate.Recordcount > 0) And (mVarOpMode = "Insert") Then
            LVField.Enabled = False
            cmd(0).Enabled = False
            lbl.Visible = True
        Else
            LVField.Enabled = True
            cmd(0).Enabled = True
            lbl.Visible = False
            rsCheck.Open "select ViewObject from [Report Modules] where NoIdx=" & cmbForm.BoundText, CNN, adOpenStatic, adLockPessimistic
            strCheck = rsCheck.Fields(0)
            strSQL = "SELECT Kolom_Object.COLUMN_NAME, Kolom_Object.DATA_TYPE From Kolom_Object Where Kolom_Object.TABLE_NAME = '" & strCheck & " ' ORDER BY Kolom_Object.COLUMN_NAME"
            Set rsField = myPart.OpenDB(strSQL)
            
            If rsField.Recordcount > 0 Then
                While Not rsField.EOF
                    Set vLst = LVField.ListItems.Add(, , UCase(rsField.Fields(0).Value))
                    vLst.SubItems(1) = UCase(rsField.Fields(1).Value)
                    rsField.MoveNext
                Wend
            End If
            
            If rsValidate.Recordcount > 0 Then
                Dim rsValidateConfig As ADODB.Recordset
                strSQL = "select * from Report_Filter where REPORT_ID = N'" & FirstNode & "'" ' and field_name = '" & Trim(LVField.ListItems(X).Text) & "'"
                Set rsValidateConfig = myPart.OpenDB(strSQL)
    
                For Y = 1 To rsValidateConfig.Recordcount
                    For x = 1 To LVField.ListItems.Count
    
                        If Trim(rsValidateConfig.Fields("FIELD_NAME").Value) = Trim(LVField.ListItems(x).Text) Then
                            LVField.ListItems(x).Checked = True
                            Exit For
                        End If
    
                    Next x
    
                    rsValidateConfig.MoveNext
                Next Y
         
            End If
            
            rsField.Close
            Set rsField = Nothing
        End If
        rsValidate.Close
        Set rsValidate = Nothing
    ElseIf (mVarOpMode = "InsertGroup") Then
            GroupAkses
    End If
   ' rsValidate.Close
    'Set rsValidate = Nothing
    Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub cmd_Click(Index As Integer)
    Dim cSave As Command
    Dim rsCheck As ADODB.Recordset
    Dim x As Integer
    Dim strComm, ft As String

    Select Case Index

        Case 0

            If mVarOpMode = "Insert" Then

                For x = 1 To LVField.ListItems.Count

                    If LVField.ListItems(x).Checked Then
                        strComm = "insert into Report_Filter (REPORT_ID,[FIELD_NAME],[FIELD_TYPE]) values (" & cmbForm.BoundText & ",'" & LVField.ListItems(x).Text & "','" & GetFieldType(Trim(LVField.ListItems(x).ListSubItems(1))) & "')"
                        SendDataToServer strComm
                    End If

                Next

                Unload Me
            ElseIf mVarOpMode = "InsertGroup" Then
                For x = 1 To LVField.ListItems.Count
                    If LVField.ListItems(x).Checked Then
                         'strComm = "insert into Report_Filter (REPORT_ID,[FIELD_NAME],[FIELD_TYPE]) values (" & cmbForm.BoundText & ",'" & LVField.ListItems(X).Text & "','" & GetFieldType(Trim(LVField.ListItems(X).ListSubItems(1))) & "')"
                         SendDataToServer (" INSERT INTO [report permit] " & _
                                                 " ([User ID], noidx, laporan,group_name)" & _
                                                 " VALUES  (" & aksess.GetID & "," & FirstNode & ", 0,'" & LVField.ListItems(x).Text & "')")
                                                
                    End If

                Next
                Unload Me
            Else

                For x = 1 To LVField.ListItems.Count
                    'delete available record
                    Set rsCheck = myPart.OpenDB("select Object_type from Report_Filter where report_id = '" & FirstNode & "' and field_name ='" & Trim(LVField.ListItems(x).Text) & "'")

                    If rsCheck.Recordcount > 0 Then
                        ft = IIf(IsNull(rsCheck.Fields(0)), "", Trim(rsCheck.Fields(0).Value))
                    Else
                        ft = ""
                    End If
               
                    strComm = "delete from Report_Filter where REPORT_ID='" & cmbForm.BoundText & "' and field_name = '" & Trim(LVField.ListItems(x).Text) & "'"
                    SendDataToServer strComm
               
                    If LVField.ListItems(x).Checked Then
                        'insert record
                        strComm = "insert into Report_Filter (report_id,field_name,field_type,object_type) values ('" & FirstNode & "','" & Trim(LVField.ListItems(x).Text) & "'," & GetFieldType(Trim(LVField.ListItems(x).ListSubItems(1).Text)) & ",'" & ft & "')"
                        SendDataToServer strComm
               
                    End If

                Next

                Unload Me
            End If

        Case 1

            Unload Me
    End Select

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)

    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Dim rsAliasRep As New DBQuick

    If mVarOpMode = "Insert" Then
        BukaComboDetailQueri cmbForm, "select [NoIdx],[Alias Report],[ViewObject] from [Report Modules]"
        cmbForm.Text = mVarFormName
'    ElseIf mVarOpMode = "InsertGroup" Then
''        Label2.Caption = "Nama Group"
''        frmEntryForm.Caption = "List Group Access Report"
'''        cmbForm.Visible = False
''        BukaComboDetailQueriGroup cmbForm, "SELECT     [group name] From [user_table_group]GROUP BY [group name]"
''        cmbForm.Text = mVarFormName
'        Label2.Visible = False
'        cmbForm.Visible = False
'        Label1.Caption = "Pilih Group Dibawah Ini Untuk Authentification Laporan"
'        GroupAkses
    Else
        BukaComboDetailQueri cmbForm, "select [NoIdx],[Alias Report],[ViewObject] from [Report Modules] where [NoIdx]='" & FirstNode & "'"
        cmbForm.Text = mVarFormName
    End If
    
'    Tengah Me
    ' Set cmbForm.RowSource = rsAliasRep.DBRecordset
End Sub

Private Sub LVField_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    If Item.Checked = False Then
        chk.Value = 0
    End If

End Sub



Private Sub GroupAkses()
Dim rsFieldGroup As ADODB.Recordset
Dim vLst As ListItem
        strSQL = "SELECT     [group name] From [user_table_group]GROUP BY [group name]"
        
       
        Set rsFieldGroup = myPart.OpenDB(strSQL)
        LVField.ColumnHeaders(1).Text = "Group Report"
        LVField.ColumnHeaders(2).Text = ""
        If rsFieldGroup.Recordcount > 0 Then
            While Not rsFieldGroup.EOF
                Set vLst = LVField.ListItems.Add(, , UCase(rsFieldGroup.Fields(0).Value))
                'vLst.SubItems(1) = UCase(rsFieldGroup.Fields(0).Value)
                rsFieldGroup.MoveNext
            Wend
        End If
End Sub


Private Sub OpenDetailReport(Param As String)
Dim sql As String
            
sql = "SELECT [report permit].noidx,[report permit].group_name,[report modules].[Alias Report]," & _
      "[report modules].ReportGroup , [report permit].Laporan " & _
      "FROM dbo.[report permit] INNER JOIN " & _
      "[report modules] ON dbo.[report permit].noidx = dbo.[report modules].NoIdx " & _
      " WHERE  (dbo.[report permit].[group_name] =" & Param & " )" & _
      " ORDER BY ReportGroup"
            
RcDetReport.DBOpen sql, CNN, lckLockBatch
'Set TDBGridGroup.DataSource = RcDetReport.DBRecordset

'TDBGrid2.Columns(0).Visible = False 'hilangkan no idx
'TDBGrid2.Columns(1).Visible = False 'hilangkan no User ID


End Sub

