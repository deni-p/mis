VERSION 5.00
Begin VB.Form FrmSetupCosting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup Costing"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10650
   DrawWidth       =   2
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSetupCosting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "FrmSetupCosting"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   10650
   Tag             =   "Setup Costing"
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00EAAF6F&
      DrawStyle       =   2  'Dot
      ForeColor       =   &H80000008&
      Height          =   5340
      Left            =   0
      ScaleHeight     =   5280
      ScaleWidth      =   10545
      TabIndex        =   0
      Top             =   0
      Width           =   10605
   End
End
Attribute VB_Name = "FrmSetupCosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RcIndirect As New DBQuick
Private RcGross As New DBQuick
Private RcPayroll As New DBQuick
Private Rcfactory As New DBQuick
Private RcMisc As New DBQuick
Private RcPartner As New DBQuick

Private Sub DgSetup_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
'If DgSetup(MenuSetup.Tab).col = 2 Then
'   DgSetup(MenuSetup.Tab).AllowUpdate = True
'Else
'   DgSetup(MenuSetup.Tab).AllowUpdate = False
'End If
End Sub

Private Sub Form_Load()
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
BomList.Tab = 0
MenuSetup.Tab = 0
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmSetupCosting
    .SetPermissions = UserAddnewDenied
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT NoItem AS [Kode Barang], ItemName AS [Nama Barang], UOM, FixCost AS [Fixed Cost], AvgCost AS [Average Cost], LastCost AS [Last Cost] FROM         Inventory WHERE     (Manufacture = 1)"
End With
Set mCall = New frmCaller
OpenDetail MyDDE.GetFieldByName("Kode Barang")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set RcIndirect = Nothing
Set RcGross = Nothing
Set RcPayroll = Nothing
Set Rcfactory = Nothing
Set RcMisc = Nothing
Set RcPartner = Nothing
Set mCall = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmSetupCosting = Nothing
End Sub

Private Sub mCall_BeforeUnload()
If FindOwnRecordset(MyDDE.ChildRecordset, "[Cost Element] = '" & MyDDE.ChildRecordset.Fields("Cost Element") & "'") = True Then
   MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Cost Element") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
   MyDDE.ChildRecordset.CancelBatch adAffectCurrent
   If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
Else
   If Not IsNull(MyDDE.ChildRecordset.Fields("Cost Element")) = True Then
      If MyDDE.ChildRecordset.Fields("Cost Element") = "" Then
         MyDDE.ChildRecordset.CancelBatch adAffectCurrent
         If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
      End If
   End If
End If
If DgSetup(MenuSetup.Tab).Enabled = True Then DgSetup(MenuSetup.Tab).SetFocus
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
With MyDDE.ChildRecordset
     .Fields(0) = mCall.GetFieldByName(0)
     .Fields(1) = mCall.GetFieldByName(1)
     .Fields(2) = 0
End With
End Sub

Private Sub MenuSetup_Click(PreviousTab As Integer)
SetDetail
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbDetail: OpenDetailPartner 1
       Case tmbSave: SaveDetail
End Select
End Sub

Private Sub PrepareQuery()
On Error Resume Next
With MyDDE
    .PrepareUpdate = " UPDATE [Inventory] Set FixCost=" & CDbl(txtBox(3)) & ",AvgCost=" & CDbl(txtBox(4)) & ",LastCost=" & CDbl(txtBox(5)) & " WHERE     ([NoItem] = N'" & txtBox(0) & "')"

    .PrepareDelete = " DELETE FROM [Inventory] WHERE   ([NoItem] = N'" & txtBox(0) & "') "
End With
Err.Clear
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbDetail:
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If MyDDE.ChildRecordset.Recordcount <> 0 Then
                  MyDDE.IsChildMemberReady = True
                  'If mAdd = True Then txtBox(3) = CDbl(LblAmount)
                  PrepareQuery
               Else
                  'MessageBox "Date detail calendar belum ada.", "Peringatan", msgOkOnly
                  MyDDE.IsChildMemberReady = False
                  PrepareQuery
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail MyDDE.GetFieldByName("Kode Barang")
End Sub


Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub SetDetail()
Select Case MenuSetup.Tab
       Case 0: Set MyDDE.ChildRecordset = RcIndirect.DBRecordset
       Case 1: Set MyDDE.ChildRecordset = RcGross.DBRecordset
       Case 2: Set MyDDE.ChildRecordset = RcPayroll.DBRecordset
       Case 3: Set MyDDE.ChildRecordset = Rcfactory.DBRecordset
       Case 4: Set MyDDE.ChildRecordset = RcMisc.DBRecordset
End Select
Set DgSetup(MenuSetup.Tab).DataSource = MyDDE.ChildRecordset
End Sub

Private Sub OpenDetail(ByVal Param As String)
RcIndirect.DBOpen "SELECT     [BOM Costing Detail].[Cost Element Type] AS [Cost Element], [Cost Element].Description AS Keterangan, [BOM Costing Detail].CostValue AS Cost FROM         [BOM Costing Detail] INNER JOIN                       [Cost Element] ON [BOM Costing Detail].[Cost Element Type] = [Cost Element].[Cost Element Type] WHERE     ([BOM Costing Detail].NoItem = N'" & Param & "') And ([BOM Costing Detail].[Group Cost] = 1) ORDER BY [BOM Costing Detail].[Cost Element Type]", CNN, lckLockBatch
RcGross.DBOpen "SELECT     [BOM Costing Detail].[Cost Element Type] AS [Cost Element], [Cost Element].Description AS Keterangan, [BOM Costing Detail].CostValue AS Cost FROM         [BOM Costing Detail] INNER JOIN                       [Cost Element] ON [BOM Costing Detail].[Cost Element Type] = [Cost Element].[Cost Element Type] WHERE     ([BOM Costing Detail].NoItem = N'" & Param & "') And ([BOM Costing Detail].[Group Cost] = 2)  ORDER BY [BOM Costing Detail].[Cost Element Type]", CNN, lckLockBatch
RcPayroll.DBOpen "SELECT     [BOM Costing Detail].[Cost Element Type] AS [Cost Element], [Cost Element].Description AS Keterangan, [BOM Costing Detail].CostValue AS Cost FROM         [BOM Costing Detail] INNER JOIN                       [Cost Element] ON [BOM Costing Detail].[Cost Element Type] = [Cost Element].[Cost Element Type] WHERE     ([BOM Costing Detail].NoItem = N'" & Param & "')  And ([BOM Costing Detail].[Group Cost] = 3)  ORDER BY [BOM Costing Detail].[Cost Element Type]", CNN, lckLockBatch
Rcfactory.DBOpen "SELECT     [BOM Costing Detail].[Cost Element Type] AS [Cost Element], [Cost Element].Description AS Keterangan, [BOM Costing Detail].CostValue AS Cost FROM         [BOM Costing Detail] INNER JOIN                       [Cost Element] ON [BOM Costing Detail].[Cost Element Type] = [Cost Element].[Cost Element Type] WHERE     ([BOM Costing Detail].NoItem = N'" & Param & "')  And ([BOM Costing Detail].[Group Cost] = 4)  ORDER BY [BOM Costing Detail].[Cost Element Type]", CNN, lckLockBatch
RcMisc.DBOpen "SELECT     [BOM Costing Detail].[Cost Element Type] AS [Cost Element], [Cost Element].Description AS Keterangan, [BOM Costing Detail].CostValue AS Cost FROM         [BOM Costing Detail] INNER JOIN                       [Cost Element] ON [BOM Costing Detail].[Cost Element Type] = [Cost Element].[Cost Element Type] WHERE     ([BOM Costing Detail].NoItem = N'" & Param & "')  And ([BOM Costing Detail].[Group Cost] = 5) ORDER BY [BOM Costing Detail].[Cost Element Type]", CNN, lckLockBatch
SetDetail
End Sub

Private Sub OpenDetailPartner(Index As Integer)
On Error GoTo Hell:
Select Case Index
       Case 1: RcPartner.DBOpen "SELECT     [Cost Element Type] AS [Cost Element], Description AS Keterangan FROM         [Cost Element] ORDER BY [Cost Element Type]", CNN, lckLockReadOnly
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 1: mCall.FromTagActive = "COST ELEMENT"
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me
Else
   MessageBox "Data Belum Ada.", "Peringatan", msgOkOnly
End If
Exit Sub
Hell:
'    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub SaveDetail()
If SendDataToServer("Delete From [BOM Costing Detail] WHERE  (NoItem = N'" & txtBox(0) & "')") = True Then
    With RcIndirect.DBRecordset
         If .Recordcount <> 0 Then
         .MoveFirst
         Do
            If .EOF Then Exit Do
             SendDataToServer " INSERT INTO [BOM Costing Detail]" & _
                              " ( NoItem, [Cost Element Type], CostValue,[Group Cost])" & _
                              " VALUES (N'" & txtBox(0) & "', N'" & .Fields("Cost Element") & "', " & .Fields("Cost") & " , 1)"
            .MoveNext
         Loop
         .MoveFirst
         End If
    End With
    
    With RcGross.DBRecordset
         If .Recordcount <> 0 Then
         .MoveFirst
         Do
            If .EOF Then Exit Do
             SendDataToServer " INSERT INTO [BOM Costing Detail]" & _
                              " ( NoItem, [Cost Element Type], CostValue,[Group Cost])" & _
                              " VALUES (N'" & txtBox(0) & "', N'" & .Fields("Cost Element") & "', " & .Fields("Cost") & " , 2)"
            .MoveNext
         Loop
         .MoveFirst
         End If
    End With
    
    With RcPayroll.DBRecordset
         If .Recordcount <> 0 Then
         .MoveFirst
         Do
            If .EOF Then Exit Do
             SendDataToServer " INSERT INTO [BOM Costing Detail]" & _
                              " ( NoItem, [Cost Element Type], CostValue,[Group Cost])" & _
                              " VALUES (N'" & txtBox(0) & "', N'" & .Fields("Cost Element") & "', " & .Fields("Cost") & " , 3)"
            .MoveNext
         Loop
         .MoveFirst
         End If
    End With
    
    With Rcfactory.DBRecordset
         If .Recordcount <> 0 Then
         .MoveFirst
         Do
            If .EOF Then Exit Do
             SendDataToServer " INSERT INTO [BOM Costing Detail]" & _
                              " ( NoItem, [Cost Element Type], CostValue,[Group Cost])" & _
                              " VALUES (N'" & txtBox(0) & "', N'" & .Fields("Cost Element") & "', " & .Fields("Cost") & " , 4)"
            .MoveNext
         Loop
         .MoveFirst
         End If
    End With
    
    With RcMisc.DBRecordset
         If .Recordcount <> 0 Then
         .MoveFirst
         Do
            If .EOF Then Exit Do
             SendDataToServer " INSERT INTO [BOM Costing Detail]" & _
                              " ( NoItem, [Cost Element Type], CostValue,[Group Cost])" & _
                              " VALUES (N'" & txtBox(0) & "', N'" & .Fields("Cost Element") & "', " & .Fields("Cost") & " , 5)"
            .MoveNext
         Loop
         .MoveFirst
         End If
    End With
End If
End Sub
