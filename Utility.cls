VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Utility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Public Enum TypeDate
'       typDateDay = 0
'       typDateMonth = 1
'       typDateYear = 2
'       typDateNow = 3
'End Enum
Private SubReportCount As Integer
Private QR1, QR2, QR3, QR4, rName1, rName2, rName3, rName4 As String
Private mVarReportQuery, mVarReportFileName, mVarReportLocation, mVarReportTitle

Private RcReport As New Recordset

Event ProgresView(ByVal Percentage As Long, ByVal GetCount As Long)

Public Property Let ReportQuery(ByVal vData As String)
  mVarReportQuery = vData
End Property
Public Function GetReportQuery() As String
   GetReportQuery = mVarReportQuery
End Function


Public Property Let ReportFileName(ByVal vData As String)
   mVarReportFileName = vData
End Property
Public Function GetReportFileName() As String
   GetReportFileName = mVarReportFileName
End Function


Public Property Let ReportLocation(ByVal vData As String)
   mVarReportLocation = vData
End Property


Public Property Let ReportTitle(ByVal vData As String)
   mVarReportTitle = vData
End Property


Public Function PendingRecord(rsTemp As ADODB.Recordset) As Boolean
On Error GoTo UnErr
Dim rsPending As New ADODB.Recordset
Set rsPending = rsTemp.Clone

rsPending.Filter = adFilterPendingRecords
If rsPending.Recordcount <> 0 Then
    PendingRecord = True
Else
    PendingRecord = False
End If
rsPending.Close
Exit Function


UnErr:
   MsgBox Err.Description, vbCritical, "ClsUtility-CheckPending"
End Function

Public Function CekDataUsed(sData As String) As Boolean
Dim rsTemp As New Recordset
Dim strSQL As String
rsTemp.CursorLocation = adUseClient
strSQL = " SELECT [Catalogue No] From [Work Order Spare Parts] WHERE ([Catalogue No] = N'" & sData & "')"
rsTemp.Open strSQL, CNN, adOpenDynamic, adLockReadOnly, adCmdText
With rsTemp
    If .Recordcount <> 0 Then
        CekDataUsed = True
    Else
        CekDataUsed = False
    End If
    .Close
End With
Set rsTemp = Nothing
End Function
Public Function CekDataSetupUsed(ByRef sTable As String, ByVal sField As String, ByVal sData As String) As Boolean
Dim rsTemp As New Recordset
Dim strSQL As String
rsTemp.CursorLocation = adUseClient
strSQL = " SELECT " & sField & " From " & sTable & " WHERE (" & sField & " = N'" & sData & "')"
rsTemp.Open strSQL, CNN, adOpenDynamic, adLockReadOnly, adCmdText
With rsTemp
    If .Recordcount <> 0 Then
        CekDataSetupUsed = True
    Else
        CekDataSetupUsed = False
    End If
    .Close
End With
Set rsTemp = Nothing
End Function
Public Function FindDoubleEntry(ByRef RecordsetName As Recordset, ByVal GetFindString As String) As Boolean
Dim mNewcari As Recordset
Set mNewcari = RecordsetName.Clone(adLockReadOnly)
With mNewcari
     If .Recordcount <> 0 Then
        .Filter = GetFindString
        If .Recordcount >= 2 Then
           FindDoubleEntry = True
        Else
           FindDoubleEntry = False
        End If
     Else
        FindDoubleEntry = False
     End If
End With
mNewcari.Close
Set mNewcari = Nothing
End Function



Public Function UpdateCost(rsData As ADODB.Recordset, nRow As Integer) As Double
Dim Avdata As Variant
Dim rsTemp As New ADODB.Recordset
Dim nCost As Double
Dim i As Integer

Set rsTemp = rsData.Clone
With rsTemp
    If .Recordcount > 0 Then
    '     Debug.Print .Source
        nCost = 0
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For i = 0 To UBound(Avdata, 2)
            nCost = nCost + IIf(IsNull(Avdata(nRow, i)), 0, Avdata(nRow, i))
        Next i
        UpdateCost = nCost
    Else
        UpdateCost = 0
    End If
End With
rsTemp.Close
Set rsTemp = Nothing
End Function
Public Sub KeyEnter(ByRef Key As Integer)
If Key = 13 Then
  SendKeys "{tab}"
End If
End Sub
Public Sub Block(ByRef obj As Object)
  obj.SelStart = 0
  obj.SelLength = Len(obj.Text)
End Sub
Public Sub Gelas(BolStatus As Boolean)
If BolStatus = True Then
   Screen.MousePointer = vbHourglass
Else
   Screen.MousePointer = vbDefault
End If
End Sub

Public Function CloseRcset(ByVal rcSetName As Recordset)
If Not rcSetName Is Nothing Then
   If rcSetName.State = 1 Then
      Set rcSetName.ActiveConnection = Nothing
      rcSetName.Close
   End If
End If
Set rcSetName = Nothing
End Function
Public Function ReadConnection() As String
On Error GoTo Hell
   Dim CON As New Connection
   CON.CursorLocation = adUseClient
    CON.ConnectionString = "Provider=SQLOLEDB.1;Password=" & _
    ";Persist Security Info=True;User ID=SA;Initial Catalog=MantellDB;Data Source=" & GetSetting("Manufacturing Intelligent", "Server", "Server Name")
   CON.Open
   If CON.State = 1 Then
      ReadConnection = CON.ConnectionString
   Else
      ReadConnection = "ERROR"
   End If
   CON.Close
   Set CON = Nothing
   Exit Function
Hell:
    ReadConnection = "ERROR"
    Set CON = Nothing
    Err.Clear
End Function

Public Sub SetDataCombo(ByVal vRowSource As ADODB.Recordset, ByRef vDataCombo As DataCombo, _
ByVal vBColumn As String, ByVal vLField As String, Optional vDataField As String)
'Set vDataCombo.DataSource = DataTrans
'vDataCombo.DataField = vDataField
Set vDataCombo.RowSource = vRowSource
vDataCombo.ListField = vLField
vDataCombo.BoundColumn = vBColumn
Set vRowSource.ActiveConnection = Nothing

End Sub
Public Function OpenDB(ByVal SelectQuery As String) As Recordset
On Error GoTo Hell
Dim Rc As New Recordset
Rc.CursorLocation = adUseClient
'Debug.Print SelectQuery
Rc.Open SelectQuery, CNN, adOpenStatic, adLockBatchOptimistic, adCmdText
Set OpenDB = Rc.Clone(adLockBatchOptimistic)
Rc.Close
Set Rc = Nothing
Exit Function
Hell:
     Set Rc = Nothing
     MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Message"
End Function
Public Function GetValue(ByVal GetSelectQuery As String) As Double
Dim Rc As New Recordset
Rc.CursorLocation = adUseClient
Rc.Open GetSelectQuery, CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
With Rc
    If .Recordcount <> 0 Then
        GetValue = .Fields(0).Value
    Else
        GetValue = 0
    End If
     .Close
End With
Set Rc = Nothing
End Function

Public Function SendCommandToServer(ByVal SelectCommand As String) As Boolean
Dim CON As New Connection
Dim Cmm As New Command
On Error GoTo Hell
With CON
     .CursorLocation = adUseClient
     .Mode = adModeShareExclusive
     .IsolationLevel = adXactIsolated
     .ConnectionString = ReadConnection
     .Open
      Set Cmm.ActiveConnection = CON
      Cmm.CommandType = adCmdText
      Cmm.CommandText = SelectCommand
'      Debug.Print SelectCommand
      Cmm.Execute
End With
SendCommandToServer = True
Set Cmm = Nothing
CON.Close
Set CON = Nothing
Exit Function
Hell:
    If CON.State = 1 Then CON.Close
'    Debug.Print SelectCommand
    Set CON = Nothing
    Set Cmm = Nothing
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Message"
    Err.Clear
End Function

Public Function PopulateView(ByVal RcName As Recordset, ByVal GetListView As ListView) As ListItem
Dim i, j As Integer
Dim Avdata As Variant
Dim Rc As New Recordset
Dim Fld As Field
Dim lst As ListItem
Dim StrTemp As String
Dim mW As Long
Rc.CursorLocation = adUseClient
Rc.Open RcName.Source, CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
With Rc
     GetListView.ColumnHeaders.Clear
     GetListView.ListItems.Clear
     GetListView.View = lvwReport
     j = 0
     For Each Fld In .Fields
         With GetListView.ColumnHeaders.Add(, , Fld.Name, GetListView.width / .Fields.Count)
              Select Case Fld.Type
                     Case 4, 3, 5, 6: If j > 0 Then .Alignment = lvwColumnRight
                     Case 11, 135: .Alignment = lvwColumnCenter
                     Case Else: .Alignment = lvwColumnLeft
              End Select
         End With
         j = j + 1
    Next Fld
    RaiseEvent ProgresView(.Recordcount, 0)
     If .Recordcount > 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        GetListView.ListItems.Clear
        j = 0
        For i = 0 To UBound(Avdata, 2)
            For Each Fld In .Fields
                If j = .Fields.Count Then j = 0
                If j = 0 Then
                   Set lst = GetListView.ListItems.Add(, , IIf(Not IsNull(Avdata(j, i)), Avdata(j, i), ""))
                Else
                   Select Case Fld.Type
                          Case 135: lst.SubItems(j) = Format(IIf(Not IsNull(Avdata(j, i)), Avdata(j, i), Date), "dd MMMM yyyy")
                          Case 11:
                               If CBool(IIf(Not IsNull(Avdata(j, i)), Avdata(j, i), False)) = True Then
                                  lst.SubItems(j) = "CLOSED"
                               Else
                                  lst.SubItems(j) = "OPEN"
                               End If
                          Case 3, 5, 6: lst.SubItems(j) = FormatNumber(IIf(Not IsNull(Avdata(j, i)), Avdata(j, i), ""), 0)
                          Case Else
                              lst.SubItems(j) = IIf(Not IsNull(Avdata(j, i)), Avdata(j, i), "")
                   End Select
                End If
                
                j = j + 1
            Next Fld
            RaiseEvent ProgresView(.Recordcount, i)
        Next i
     End If
     .Close
End With
Set PopulateView = lst
Set Rc = Nothing
Set Fld = Nothing
Set lst = Nothing
Set Avdata = Nothing
End Function

Public Function PopulateTreeView(ByVal RcName As Recordset, ByVal GetTreeView As TreeView, ByVal NodeKey As String) As Node
Dim i As Integer
Dim Avdata As Variant
Dim Rc As New Recordset
Dim lst As Node
Set Rc = RcName.Clone(adLockReadOnly)
With Rc
     If .Recordcount > 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For i = 0 To UBound(Avdata, 2)
            Set lst = GetTreeView.Nodes.Add(NodeKey, tvwChild, "KEY" & IIf(Not IsNull(Avdata(0, i)), Avdata(0, i), ""), IIf(Not IsNull(Avdata(0, i)), Avdata(0, i), ""))
            lst.EnsureVisible
            lst.Expanded = True
        Next i
     End If
     .Close
End With
Set PopulateTreeView = lst
Set Rc = Nothing
Set lst = Nothing
Set Avdata = Nothing
End Function

Public Function PopulateGrid(ByVal RcName As Recordset) As Recordset
Dim i As Integer
Dim Avdata As Variant
Dim Rc As New Recordset
Dim Fld As Field
'Set Rc = RcName.Clone(adLockReadOnly)
With Rc
     i = 0
     For Each Fld In RcName.Fields
'         MsgBox fld.Value
         Rc.Fields.Append Fld.Name, adBSTR
'         GetDataGrid.Columns.Add I
'         GetDataGrid.Columns(I).Caption = fld.Name
'         GetDataGrid.Columns(I).DataField = fld.Name
'         I = I + 1
     Next
     Set PopulateGrid = Rc
'     If .RecordCount > 0 Then
'        Avdata = .GetRows(.RecordCount, adBookmarkFirst)
'        For I = 0 To UBound(Avdata, 2)
'        Next I
'     End If
'     .Close
End With
Set Rc = Nothing
Set Avdata = Nothing
End Function

Public Function SimpleAutoNo(ByVal SelectQuery As String) As String
Dim Rc As New Recordset
Rc.CursorLocation = adUseClient
Rc.Open SelectQuery, CNN, adOpenStatic, adLockBatchOptimistic, adCmdText
'Debug.Print SelectQuery
With Rc
     If .Recordcount <> 0 Then
        SimpleAutoNo = Trim(Str(IIf(Not IsNull(.Fields(0)), .Fields(0), Year(Date))))
     Else
        SimpleAutoNo = "1"
     End If
End With
Rc.Close
Set Rc = Nothing
End Function

Public Function IsFounded(ByVal GetSelectQuery As String) As Boolean
Dim Rc As New Recordset
Rc.CursorLocation = adUseClient
Rc.Open GetSelectQuery, CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
With Rc
     If .Recordcount > 0 Then
        IsFounded = True
     End If
     .Close
End With
Set Rc = Nothing
End Function

'Public Function DefaultDateTime(ByVal GetTypeDate As TypeDate, Optional ByVal ParamDate As Date) As Variant
'Dim sTmp As Date
'sTmp = ParamDate
'Select Case GetTypeDate
'       Case typDateDay:
'            DefaultDateTime = Day(sTmp)
'       Case typDateMonth:
'            DefaultDateTime = Month(sTmp)
'       Case typDateYear:
'            DefaultDateTime = Year(sTmp)
'       Case typDateNow:
'            DefaultDateTime = sTmp
'End Select
'End Function

Public Function DaysInMonth(ByVal d As Date) As Long
  Select Case Month(d)
    Case 2
      If LeapYear(Year(d)) Then
        DaysInMonth = 29
      Else
        DaysInMonth = 28
      End If
    Case 4, 6, 9, 11
      DaysInMonth = 30
    Case 1, 3, 5, 7, 8, 10, 12
      DaysInMonth = 31
  End Select
End Function

Public Function LeapYear(ByVal YYYY As Long) As Boolean
        LeapYear = YYYY Mod 4 = 0 _
                   And (YYYY Mod 100 <> 0 Or YYYY Mod 400 = 0)
End Function

Public Sub CallReportView(ByVal SelectQuery As String, ByVal ReportName As String, _
ByVal ReportFilePath As String, Optional ByVal TitleReport As String, Optional ReportParams As String)  ', Optional ParPeriod As String)

On Error GoTo Hell
'Debug.Print TitleReport
Screen.MousePointer = vbHourglass
   Dim Mprint As New frmReportView
   With Mprint
'      Debug.Print SelectQuery
      If rName1 <> "" Then
        .SubReport1 QR1, rName1
      End If
      If rName2 <> "" Then
        .SubReport2 QR2, rName2
      End If
      If rName3 <> "" Then
        .SubReport3 QR3, rName3
      End If
      If rName4 <> "" Then
        .SubReport4 QR4, rName4
      End If

      .CallReport SelectQuery, ReportName, ReportFilePath, TitleReport, ReportParams
      Screen.MousePointer = vbDefault
      .Show
'      FrmClosing.SetFocus
   End With
   
   QR1 = ""
   QR2 = ""
   QR3 = ""
   QR4 = ""
   
   rName1 = ""
   rName2 = ""
   rName3 = ""
   rName4 = ""
   Screen.MousePointer = vbDefault
Exit Sub
Hell:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Message"
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub

Public Sub CallReportViewer()
  CallReportView mVarReportQuery, mVarReportFileName, mVarReportLocation, mVarReportTitle
End Sub

Public Sub SendKoman(ByVal SelectCommand As String)
Dim CON As New Connection
Dim Cmm As New Command
On Error GoTo Hell
With CON
     .CursorLocation = adUseClient
     .Mode = adModeShareExclusive
     .IsolationLevel = adXactIsolated
     .ConnectionString = ReadConnection
     .Open
      Set Cmm.ActiveConnection = CON
      Cmm.CommandType = adCmdText
      Cmm.CommandText = SelectCommand
'      Debug.Print SelectCommand
      Cmm.Execute
End With
'Debug.Print CON
'Debug.Print SelectCommand
Set Cmm = Nothing
CON.Close
Set CON = Nothing
Exit Sub
Hell:
    If CON.State = 1 Then CON.Close
'    Debug.Print SelectCommand
    Set CON = Nothing
    Set Cmm = Nothing
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Message"
'    Debug.Print SelectCommand
    Err.Clear
End Sub

Public Sub AddSubReport(ByVal SelectQuery As String, ByVal ReportName As String)
   SubReportCount = SubReportCount + 1
   If SubReportCount = 5 Then
     SubReportCount = 1
   End If
   
   Select Case SubReportCount
      Case 1
         QR1 = SelectQuery
         rName1 = ReportName
      Case 2
        QR2 = SelectQuery
        rName2 = ReportName
      Case 3
        QR3 = SelectQuery
        rName3 = ReportName
      Case 4
        QR4 = SelectQuery
        rName4 = ReportName
   End Select
End Sub


Public Function MayDo(FormName) As Boolean
   Dim rs As ADODB.Recordset
   Dim strSQL As String
   strSQL = "select edit from privilege where [User Name] ='" & _
            Trim(MainMenu.StatusBar1.Panels(1).Text) & "' and [Form List] = '" & Trim(FormName) & "'"
   
   Set rs = OpenDB(strSQL)
   If rs.Recordcount > 0 Then
      MayDo = rs.Fields(0).Value
   Else
      MayDo = False
   End If
End Function

Public Sub SelectGrid(GridObj As DataGrid, _
                      ByVal xCol As Integer, _
                      ByVal yRow As Integer)

    If yRow = -1 And xCol = -1 Then
        GridObj.SelStartCol = 0
        GridObj.SelEndCol = GridObj.Columns.Count
    End If

End Sub


