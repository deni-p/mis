VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsRecursive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mOrderQTY As Long
Private StrQtyBom, mVarBomFirst As String
Private mVarBom As Boolean
Private Rc As New DBQuick
Private RcDetail As New DBQuick
Private RcDummy As New DBQuick
Private mVarLoadBom, mVarNextBom As Boolean
Private mVarOldPointer, mNext, mMove As Long

Public Sub ReadComponentOrder(ByVal HorizontalDays As Date, ByVal PlanningDays As Long, ByVal ObjProg As ProgressBar)
On Error GoTo Hell
'Rc.DBOpen " SELECT [Ord Comp Detail].StageID AS SeqStageID, [Ord Comp Detail].NoItem AS Component, [Ord Comp Detail].[DESC] AS [Nama Component],  [Ord Comp Detail].UOM, [Ord Comp Detail].PartnerID, [Ord Comp Detail].[Quote Qty] AS [QTY Usage], Inventory.Manufacture, [Manufacture Order].NoItem AS [BOM ID]" & _
          " FROM [Ord Comp Detail] INNER JOIN Inventory ON [Ord Comp Detail].NoItem = Inventory.NoItem INNER JOIN [Manufacture Order] ON [Ord Comp Detail].OrderID = [Manufacture Order].OrderID ORDER BY Inventory.Manufacture, [Ord Comp Detail].NoItem", Cnn
'Rc.DBOpen " Shape{SELECT [Manufacture Order].NoItem, [Manufacture Order].OrderName, Inventory.Manufacture, [Ord Comp Detail].[Quote Qty] AS [Quote qty], Inventory.LeadTimeDays, Inventory.PartnerID, Inventory.MinStock AS [Min Stock], Inventory.ROP AS RQTY,  [Order Output Detail].StartDate - 1 AS RequireDate, [Manufacture Order].[QTY Order], [Manufacture Order].OrderID, [Manufacture Order].Note, Inventory.BomReff" & _
          "       FROM Inventory INNER JOIN [Ord Comp Detail] INNER JOIN [Order Output Detail] INNER JOIN [Manufacture Order] ON [Order Output Detail].OrderID = [Manufacture Order].OrderID ON [Ord Comp Detail].OrderID = [Order Output Detail].OrderID AND [Ord Comp Detail].StageID = [Order Output Detail].StageID ON Inventory.NoItem = [Manufacture Order].NoItem WHERE  ([Ord Comp Detail].Complete = 0) AND ([Order Output Detail].StartDate <= CONVERT(DATETIME, '" & Format(HorizontalDays, "dd/mm/yy") & "', 3)) GROUP BY [Manufacture Order].NoItem, [Manufacture Order].OrderName, [Ord Comp Detail].[Quote Qty], Inventory.LeadTimeDays, Inventory.PartnerID,  Inventory.MinStock, Inventory.ROP, [Order Output Detail].StartDate - 1, Inventory.Manufacture, [Manufacture Order].[QTY Order],  [Manufacture Order].OrderID, [Manufacture Order].Note, Inventory.BomReff ORDER BY [Manufacture Order].OrderID, [Manufacture Order].NoItem} As ParentMenu" & _
          " Append({SELECT [BOM Component Detail].SeqStageID, [BOM Component Detail].Component AS [cOMPONENT], Inventory.ItemName AS [Nama Komponen],  [BOM Component Detail].UOM, Inventory.PartnerID AS [Partner ID], PartnerDB.CompanyName AS [Nama Perusahaan], [BOM Component Detail].QTYUsage,Inventory.Manufacture,[BOM Component Detail].NoItem FROM  [BOM Component Detail] INNER JOIN [BOM Stage Detail] ON [BOM Component Detail].SeqStageID = [BOM Stage Detail].SeqStageID AND [BOM Component Detail].NoItem = [BOM Stage Detail].NoItem AND [BOM Component Detail].BomReff = [BOM Stage Detail].BomReff INNER JOIN" & _
          " Inventory INNER JOIN PartnerDB ON Inventory.PartnerID = PartnerDB.PartnerID ON [BOM Component Detail].BomReff = Inventory.BomReff AND  [BOM Component Detail].Component = Inventory.NoItem  GROUP BY [BOM Component Detail].SeqStageID, [BOM Component Detail].Component, Inventory.ItemName, [BOM Component Detail].UOM, Inventory.PartnerID,  PartnerDB.CompanyName, [BOM Component Detail].QTYUsage, [BOM Stage Detail].NoLine,Inventory.Manufacture,[BOM Component Detail].NoItem ORDER BY Inventory.Manufacture,[BOM Stage Detail].NoLine, [BOM Component Detail].SeqStageID}as ChildMd Relate NoItem to NoItem)", Cnn
    
Rc.DBOpen " SELECT  [Ord Comp Detail].NoItem, Inventory.ItemName, Inventory.Manufacture, [Ord Comp Detail].[Quote Qty] AS [Quote qty], Inventory.LeadTimeDays,  Inventory.PartnerID, Inventory.MinStock AS [Min Stock], Inventory.ROP AS RQTY, [Order Output Detail].StartDate - 1 AS RequireDate,  [Manufacture Order].[QTY Order], [Manufacture Order].OrderID, [Manufacture Order].Note, [Order Output Detail].StartDate,  [Order Output Detail].EndDate" & _
          " FROM [Order Output Detail] INNER JOIN [Manufacture Order] ON [Order Output Detail].OrderID = [Manufacture Order].OrderID INNER JOIN [Ord Comp Detail] INNER JOIN Inventory ON [Ord Comp Detail].NoItem = Inventory.NoItem ON [Order Output Detail].OrderID = [Ord Comp Detail].OrderID " & _
          " WHERE ([Ord Comp Detail].Complete = 0) AND ([Manufacture Order].CreateDate <= CONVERT(DATETIME, '" & Format(HorizontalDays, "dd/mm/yy") & "', 3)) GROUP BY [Ord Comp Detail].NoItem, Inventory.ItemName, [Ord Comp Detail].[Quote Qty], Inventory.LeadTimeDays, Inventory.PartnerID, Inventory.MinStock,  Inventory.ROP, [Order Output Detail].StartDate - 1, Inventory.Manufacture, [Manufacture Order].[QTY Order], [Manufacture Order].OrderID,  [Manufacture Order].Note, [Order Output Detail].StartDate, [Order Output Detail].EndDate ORDER BY Inventory.Manufacture, [Ord Comp Detail].NoItem", CNN
'    Debug.Print Rc.DBRecordset.Source
mVarOldPointer = 0
mVarLoadBom = False
mMove = 1

Set RcDummy.DBRecordset = Rc.DBRecordset.Clone(adLockReadOnly)
mVarBomFirst = ""
SendDataToServer ("DELETE FROM [Planned Order]")
With Rc.DBRecordset
     If .Recordcount <> 0 Then
     ObjProg.Max = .Recordcount
     ObjProg.Min = 0
     Do
       If .EOF Then Exit Do
       mVarBomFirst = mVarBomFirst & vbCrLf
       If .Fields(2) = True Then
          CallSimpanData .Fields(0), .Fields("StartDate"), .Fields(9), PlanningDays, .Fields(2), True
          LinkManufacture .Fields(0)
          If RcDummy.DBRecordset.Recordcount <> 0 Then
             mOrderQTY = .Fields(9)
             RecursiveCallManufacture RcDummy.DBRecordset, .AbsolutePosition, .Fields("StartDate"), PlanningDays, .Fields(9)
          End If
         'MessageBox mVarBomFirst
       Else
         CallSimpanData .Fields(0), .Fields("StartDate"), .Fields(9), PlanningDays, .Fields(2)
       End If
       ObjProg.value = ObjProg.value + 1
       .MoveNext
       
     Loop
     End If
     'MessageBox mVarBomFirst
End With
SendDataToServer ("DELETE FROM [Planned Order] WHERE     ([Suggest QTY] = 0)")
'Dim Rc As New DBQuick
'Dim k As Long
'Dim Avdata, mVarQTYPO, mVarStock, mVarSug As Variant
'Rc.DBOpen " Shape{SELECT [Manufacture Order].NoItem, [Manufacture Order].OrderName, Inventory.Manufacture, [Ord Comp Detail].[Quote Qty] AS [Quote qty], Inventory.LeadTimeDays, Inventory.PartnerID, Inventory.MinStock AS [Min Stock], Inventory.ROP AS RQTY,  [Order Output Detail].StartDate - 1 AS RequireDate, [Manufacture Order].[QTY Order], [Manufacture Order].OrderID, [Manufacture Order].Note, Inventory.BomReff" & _
'          "       FROM Inventory INNER JOIN [Ord Comp Detail] INNER JOIN [Order Output Detail] INNER JOIN [Manufacture Order] ON [Order Output Detail].OrderID = [Manufacture Order].OrderID ON [Ord Comp Detail].OrderID = [Order Output Detail].OrderID AND [Ord Comp Detail].StageID = [Order Output Detail].StageID ON Inventory.NoItem = [Manufacture Order].NoItem WHERE  ([Ord Comp Detail].Complete = 0) AND ([Order Output Detail].StartDate <= CONVERT(DATETIME, '" & Format(HorizontalDays, "dd/mm/yy") & "', 3)) GROUP BY [Manufacture Order].NoItem, [Manufacture Order].OrderName, [Ord Comp Detail].[Quote Qty], Inventory.LeadTimeDays, Inventory.PartnerID,  Inventory.MinStock, Inventory.ROP, [Order Output Detail].StartDate - 1, Inventory.Manufacture, [Manufacture Order].[QTY Order],  [Manufacture Order].OrderID, [Manufacture Order].Note, Inventory.BomReff ORDER BY [Manufacture Order].OrderID, [Manufacture Order].NoItem} As ParentMenu" & _
'          " Append({SELECT [BOM Component Detail].SeqStageID, [BOM Component Detail].Component AS [cOMPONENT], Inventory.ItemName AS [Nama Komponen],  [BOM Component Detail].UOM, Inventory.PartnerID AS [Partner ID], PartnerDB.CompanyName AS [Nama Perusahaan], [BOM Component Detail].QTYUsage,Inventory.Manufacture,[BOM Component Detail].NoItem FROM  [BOM Component Detail] INNER JOIN [BOM Stage Detail] ON [BOM Component Detail].SeqStageID = [BOM Stage Detail].SeqStageID AND [BOM Component Detail].NoItem = [BOM Stage Detail].NoItem AND [BOM Component Detail].BomReff = [BOM Stage Detail].BomReff INNER JOIN" & _
'          " Inventory INNER JOIN PartnerDB ON Inventory.PartnerID = PartnerDB.PartnerID ON [BOM Component Detail].BomReff = Inventory.BomReff AND  [BOM Component Detail].Component = Inventory.NoItem  GROUP BY [BOM Component Detail].SeqStageID, [BOM Component Detail].Component, Inventory.ItemName, [BOM Component Detail].UOM, Inventory.PartnerID,  PartnerDB.CompanyName, [BOM Component Detail].QTYUsage, [BOM Stage Detail].NoLine,Inventory.Manufacture,[BOM Component Detail].NoItem ORDER BY Inventory.Manufacture,[BOM Stage Detail].NoLine, [BOM Component Detail].SeqStageID} Relate NoItem to NoItem)", Cnn
'With Rc.DBRecordset
'     If .Recordcount <> 0 Then
'        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
'        SendDataToServer ("DELETE FROM [Planned Order]")
'        For k = 0 To UBound(Avdata, 2)
'            If CBool(Avdata(2, k)) = True Then RecursiveCallManufacture Avdata(0, k)
''             mOrderQTY = Avdata(9, k)
''             mVarStock = CekStock(Avdata(0, k))
''             If CDbl(Avdata(3, k)) > mVarStock Then
''                mVarQTYPO = CDbl(Avdata(3, k)) - mVarStock
''                If mVarQTYPO < 0 Then mVarQTYPO = mVarQTYPO * (-1)
''                If mVarQTYPO >= CDbl(Avdata(6, k)) Then
''                   mVarSug = mVarQTYPO + CDbl(Avdata(7, k))
'''                  mVarSug = mVarQTYPO + CDbl(Avdata(3, k)) + CDbl(Avdata(7, k))
''                Else
''                   mVarSug = mVarQTYPO
''                End If
''                  'Tgl Required + lead
'''                  SendDataToServer (" INSERT INTO [Planned Order]" & _
'''                                    " (OrderID,Note,NoItem, [DESC], M_OR_P, PartnerID, [Suggest QTY], [Order QTY], [Required Date], [Suggest Order Date], [Order Date])" & _
'''                                    " VALUES   (N'" & Avdata(10, k) & "',N'" & Avdata(11, k) & "',N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(2, k)) & ", N'" & Avdata(5, k) & "', " & CDbl(mVarSug) & ", " & CDbl(mVarSug) & ", CONVERT(DATETIME, '" & Format(Avdata(8, k), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format((CDbl(Avdata(8, k)) - CDbl(Avdata(4, k)) + CDbl(Text1)), "dd/mm/yy") & "', 3), CONVERT(DATETIME,'" & Format(Avdata(8, k) + CDbl(Avdata(4, k)), "dd/mm/yy") & "', 3))")
'''               Else
'''                  SendDataToServer (" INSERT INTO [Planned Order]" & _
'''                                    " (OrderID,Note,NoItem, [DESC], M_OR_P, PartnerID, [Suggest QTY], [Order QTY], [Required Date], [Suggest Order Date], [Order Date])" & _
'''                                    " VALUES   (N'" & Avdata(10, k) & "',N'" & Avdata(11, k) & "',N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(2, k)) & ", N'" & Avdata(5, k) & "', " & CDbl(Avdata(3, k)) & ", 0, CONVERT(DATETIME, '" & Format(Avdata(8, k), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format((CDbl(Avdata(8, k)) - CDbl(Avdata(4, k)) + CDbl(Text1)), "dd/mm/yy") & "', 3), CONVERT(DATETIME,'" & Format(CDbl(Avdata(8, k)) + CDbl(Avdata(4, k)), "dd/mm/yy") & "', 3))")
''
''                  SendDataToServer (" INSERT INTO [Planned Order]" & _
''                                    " (OrderID,Note,NoItem, [DESC], M_OR_P, PartnerID, [Suggest QTY], [Order QTY], [Required Date], [Suggest Order Date], [Order Date])" & _
''                                    " VALUES   (N'" & Avdata(10, k) & "',N'" & Avdata(11, k) & "',N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(2, k)) & ", N'" & Avdata(5, k) & "', " & CDbl(mVarSug) & ", " & CDbl(mVarSug) & ", CONVERT(DATETIME, '" & Format(Avdata(8, k), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format((CDbl(Avdata(8, k)) - CDbl(Avdata(4, k)) + CDbl(PlanningDays)), "dd/mm/yy") & "', 3), CONVERT(DATETIME,'" & Format((CDbl(Avdata(8, k)) - CDbl(Avdata(4, k)) + CDbl(PlanningDays)), "dd/mm/yy") & "', 3))")
''               Else
''                  SendDataToServer (" INSERT INTO [Planned Order]" & _
''                                    " (OrderID,Note,NoItem, [DESC], M_OR_P, PartnerID, [Suggest QTY], [Order QTY], [Required Date], [Suggest Order Date], [Order Date])" & _
''                                    " VALUES   (N'" & Avdata(10, k) & "',N'" & Avdata(11, k) & "',N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(2, k)) & ", N'" & Avdata(5, k) & "', " & CDbl(Avdata(3, k)) & ", 0, CONVERT(DATETIME, '" & Format(Avdata(8, k), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format((CDbl(Avdata(8, k)) - CDbl(Avdata(4, k)) + CDbl(PlanningDays)), "dd/mm/yy") & "', 3), CONVERT(DATETIME,'" & Format((CDbl(Avdata(8, k)) - CDbl(Avdata(4, k)) + CDbl(PlanningDays)), "dd/mm/yy") & "', 3))")
''
''               End If
''            If CBool(Avdata(2, k)) = False Then
''               ReadBOMBOM Avdata(0, k), Avdata(8, k), Avdata(9, k), PlanningDays
''            End If
'        Next
'
'     End If
'End With
'Set Avdata = Nothing
Hell:
    Err.Clear
End Sub

Private Sub RecursiveCallManufacture(ByVal RecData As Recordset, ByVal Index As Long, ByVal HorizontalDays As Date, ByVal PlanningDays As Long, ByVal QtyOrder As Long)
Dim Avdata As Variant
Dim I As Long
'SendDataToServer
'     0          1             2          3      4           5            6         7
'SeqStageID  Component   Nama Component  UOM PartnerID   QTY Usage   Manufacture BOM ID
If RecData.EOF Then Exit Sub
mVarBomFirst = mVarBomFirst & "," & RecData.Fields(1)
CallSimpanData RecData.Fields(1), HorizontalDays, QtyOrder, PlanningDays, True
With RecData
     'MsgBox .Source
     If .Recordcount <> 0 Then
        If .Fields(6) = True Then
           If .AbsolutePosition <> .Recordcount Then
              If Not .EOF Then
                 .AbsolutePosition = .AbsolutePosition + 1
                 RecursiveCallManufacture RecData, .AbsolutePosition, HorizontalDays, PlanningDays, QtyOrder
              End If
           Else
              Exit Sub
           End If
        Else
           .AbsolutePosition = .AbsolutePosition + 1
            RecursiveCallManufacture RecData, .AbsolutePosition, HorizontalDays, PlanningDays, QtyOrder
        End If
     End If
End With
End Sub

Private Sub LinkManufacture(ByVal ParamBOM As String)
RcDetail.DBOpen " SELECT     [BOM Component Detail].SeqStageID, [BOM Component Detail].Component, [BOM Component Detail].Description, [BOM Component Detail].UOM,  Inventory.PartnerID, [BOM Component Detail].QTYUsage AS [QTY Usage], Inventory.Manufacture FROM [BOM Component Detail] INNER JOIN Inventory ON [BOM Component Detail].Component = Inventory.NoItem WHERE     ([BOM Component Detail].NoItem = N'" & ParamBOM & "') ORDER BY Inventory.Manufacture", CNN, lckLockReadOnly
With RcDetail.DBRecordset
     Set RcDummy.DBRecordset = RcDetail.DBRecordset.Clone(adLockReadOnly)
     'If .Recordcount <> 0 Then RecursiveCallManufacture RcDummy.DBRecordset, .AbsolutePosition
        
End With
RcDetail.CloseDB
Set RcDetail = Nothing
End Sub

Private Function CariManufacture(ByVal Rec As Recordset, ByVal No_Item As String, Optional MoveRec As Long)
If Rec.Recordcount <> 0 Then
   If MoveRec = 0 Then MoveRec = Rec.AbsolutePosition
   Rec.AbsolutePosition = MoveRec
   If Not Rec.EOF Then
      MoveRec = MoveRec + 1
      CariManufacture = CariManufacture(Rec, Rec.Fields(1), MoveRec)
   End If
End If
End Function

Private Sub CallSimpanData(ByVal Param As String, ByVal TglRequireDate As String, ByVal QtyOrder As Long, ByVal PlanningDays As Long, ByVal M_OR_P As Boolean, Optional ByVal HeaderManufacture As Boolean = False)
Dim RcBOM As New DBQuick
Dim Avdata As Variant
Dim mVarQTYPO As Variant
Dim mVarStock As Variant
Dim mVarStockTemp As Variant
Dim mVarSug As Variant
Dim mParam As String
Dim k As Integer
If M_OR_P = True Then
   If HeaderManufacture = True Then
      RcBOM.DBOpen " SELECT     [Ord Comp Detail].NoItem, [Ord Comp Detail].[DESC] AS ItemName, Inventory.PartnerID, Inventory.ROP AS RQTY, Inventory.MinStock AS [Min Stock],  [Ord Comp Detail].[Quote Qty] AS [Quote QTY], Inventory.LeadTimeDays, Inventory.Manufacture FROM         Inventory INNER JOIN   [Ord Comp Detail] ON Inventory.NoItem = [Ord Comp Detail].NoItem GROUP BY Inventory.LeadTimeDays, Inventory.PartnerID, Inventory.ROP, Inventory.MinStock, Inventory.Manufacture, [Ord Comp Detail].NoItem,   [Ord Comp Detail].[DESC], [Ord Comp Detail].[Quote Qty] HAVING      ([Ord Comp Detail].NoItem = N'" & Param & "')", CNN
   Else
      RcBOM.DBOpen " SELECT   [BOM Component Detail].Component, Inventory.ItemName, Inventory.PartnerID, Inventory.ROP AS RQTY,  Inventory.MinStock AS [Min Stock], [BOM Component Detail].QTYUsage AS [Quote QTY], Inventory.LeadTimeDays, Inventory.Manufacture FROM         [BOM Component Detail] INNER JOIN Inventory ON [BOM Component Detail].Component = Inventory.NoItem WHERE     ([BOM Component Detail].NoItem = N'" & Param & "') GROUP BY Inventory.LeadTimeDays, Inventory.PartnerID, Inventory.ROP, Inventory.ItemName, [BOM Component Detail].QTYUsage, Inventory.MinStock,  Inventory.Manufacture, [BOM Component Detail].Component ORDER BY [BOM Component Detail].Component", CNN
   End If
Else
   RcBOM.DBOpen " SELECT [Ord Comp Detail].NoItem, Inventory.ItemName, Inventory.PartnerID, Inventory.ROP AS RQTY, Inventory.MinStock AS [Min Stock],  [Ord Comp Detail].[Quote Qty] AS [Quote QTY], Inventory.LeadTimeDays, Inventory.Manufacture FROM Inventory INNER JOIN  [Ord Comp Detail] ON Inventory.NoItem = [Ord Comp Detail].NoItem INNER JOIN    [Manufacture Order] ON [Ord Comp Detail].OrderID = [Manufacture Order].OrderID WHERE     ([Ord Comp Detail].NoItem = N'" & Param & "') GROUP BY Inventory.LeadTimeDays, Inventory.PartnerID, Inventory.ROP, [Ord Comp Detail].NoItem, Inventory.ItemName, Inventory.MinStock,  Inventory.Manufacture, [Ord Comp Detail].[Quote Qty] HAVING      (Inventory.Manufacture = 0) ORDER BY [Ord Comp Detail].NoItem", CNN
End If
With RcBOM.DBRecordset

     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For k = 0 To UBound(Avdata, 2)
            If CBool(Avdata(7, k)) = True Then
               mVarStock = CekStock(Avdata(0, k))
               If CDbl(Avdata(5, k)) > mVarStock Then
                  mVarQTYPO = CDbl(Avdata(5, k)) - mVarStock
                  If mVarQTYPO < 0 Then mVarQTYPO = mVarQTYPO * (-1)
                  If mVarQTYPO >= CDbl(Avdata(4, k)) Then
                     If CBool(Avdata(7, k)) = False Then
                        mVarSug = (CDbl(Avdata(5, k)) + CDbl(Avdata(3, k)))
                     Else
                        mVarSug = CDbl(Avdata(5, k)) '+ CDbl(Avdata(3, k)))
                     End If
                  Else
                     mVarSug = mVarQTYPO
                  End If
                  SendDataToServer (" INSERT INTO [Planned Order]" & _
                                    " ([Quote Qty],NoItem, [DESC], M_OR_P, PartnerID, [Suggest QTY], [Order QTY], [Required Date], [Suggest Order Date], [Order Date])" & _
                                    " VALUES   (" & CDbl(Avdata(5, k)) & ",N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(7, k)) & ", N'" & Avdata(2, k) & "', " & mVarSug & ", " & mVarSug & ", CONVERT(DATETIME, '" & Format(CDate(TglRequireDate), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) - CDbl(PlanningDays)), "dd/mm/yy") & "', 3), CONVERT(DATETIME,'" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) - CDbl(PlanningDays)), "dd/mm/yy") & "', 3))")
                  
               
               Else
                    SendDataToServer (" INSERT INTO [Planned Order]" & _
                                      " ([Quote Qty],NoItem, [DESC], M_OR_P, PartnerID, [Suggest QTY], [Order QTY], [Required Date], [Suggest Order Date], [Order Date])" & _
                                      " VALUES   (" & CDbl(Avdata(5, k)) & ",N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(7, k)) & ", N'" & Avdata(2, k) & "', " & CDbl(Avdata(5, k)) & ", " & CDbl(Avdata(5, k)) & ", CONVERT(DATETIME, '" & Format(CDate(TglRequireDate), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) - CDbl(PlanningDays)), "dd/mm/yy") & "', 3), CONVERT(DATETIME,'" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) - CDbl(PlanningDays)), "dd/mm/yy") & "', 3))")
               
               End If
               
               Call ReadDetailBOMBOM(Avdata(0, k), TglRequireDate, QtyOrder, PlanningDays)
            Else
               mVarStock = CekStock(Avdata(0, k))
               If CDbl(Avdata(5, k)) > mVarStock Then
                  mVarQTYPO = CDbl(Avdata(5, k)) - mVarStock
                  If mVarQTYPO < 0 Then mVarQTYPO = mVarQTYPO * (-1)
                  If mVarQTYPO >= CDbl(Avdata(4, k)) Then
                     If CBool(Avdata(7, k)) = False Then
                        mVarSug = (CDbl(Avdata(5, k)) + CDbl(Avdata(3, k)))
                     Else
                        mVarSug = CDbl(Avdata(5, k)) '+ CDbl(Avdata(3, k)))
                     End If
                  Else
                     mVarSug = mVarQTYPO
                  End If
                  SendDataToServer (" INSERT INTO [Planned Order]" & _
                                    " ([Quote Qty],NoItem, [DESC], M_OR_P, PartnerID, [Suggest QTY], [Order QTY], [Required Date], [Suggest Order Date], [Order Date])" & _
                                    " VALUES   (" & CDbl(Avdata(5, k)) & ",N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(7, k)) & ", N'" & Avdata(2, k) & "', " & mVarSug & ", " & mVarSug & ", CONVERT(DATETIME, '" & Format(CDate(TglRequireDate), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) - CDbl(PlanningDays)), "dd/mm/yy") & "', 3), CONVERT(DATETIME,'" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) - CDbl(PlanningDays)), "dd/mm/yy") & "', 3))")
                  
               
               Else
                  mVarQTYPO = CDbl(Avdata(5, k)) ' - mVarStock
                  If mVarQTYPO < 0 Then mVarQTYPO = mVarQTYPO * (-1)
                  If mVarQTYPO >= CDbl(Avdata(4, k)) Then
                     If CBool(Avdata(7, k)) = False Then
                        mVarSug = (CDbl(Avdata(5, k)) + CDbl(Avdata(3, k)))
                     Else
                        mVarSug = CDbl(Avdata(5, k)) '+ CDbl(Avdata(3, k)))
                     End If
                  Else
                     mVarSug = mVarQTYPO
                  End If
                    SendDataToServer (" INSERT INTO [Planned Order]" & _
                                      " ([Quote Qty],NoItem, [DESC], M_OR_P, PartnerID, [Suggest QTY], [Order QTY], [Required Date], [Suggest Order Date], [Order Date])" & _
                                      " VALUES   (" & CDbl(Avdata(5, k)) & ",N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(7, k)) & ", N'" & Avdata(2, k) & "', " & CDbl(mVarSug) & ", " & CDbl(mVarSug) & ", CONVERT(DATETIME, '" & Format(CDate(TglRequireDate), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) - CDbl(PlanningDays)), "dd/mm/yy") & "', 3), CONVERT(DATETIME,'" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) - CDbl(PlanningDays)), "dd/mm/yy") & "', 3))")
               
               End If
            End If
        Next
     End If
End With
Set Avdata = Nothing
End Sub

Private Sub ReadBOMBOM(ByVal Param As String, ByVal TglRequireDate As String, ByVal QtyOrder As Long, ByVal PlanningDays As Long)
Dim RcBOM As New DBQuick
Dim Avdata As Variant
Dim mVarQTYPO As Variant
Dim mVarStock As Variant
Dim mVarSug As Variant
Dim mParam As String
Dim k As Integer
RcBOM.DBOpen " SELECT [BOM Component Detail].NoItem, Inventory.ItemName, Inventory.PartnerID, Inventory.ROP AS RQTY, Inventory.MinStock AS [Min Stock],  [BOM Component Detail].QTYUsage AS [Quote QTY], Inventory.LeadTimeDays, Inventory.Manufacture  FROM         [BOM Component Detail] INNER JOIN                       Inventory ON [BOM Component Detail].NoItem = Inventory.NoItem WHERE     ([BOM Component Detail].Component = N'" & Param & "') GROUP BY Inventory.LeadTimeDays, Inventory.PartnerID, Inventory.ROP, [BOM Component Detail].NoItem, Inventory.ItemName,                        [BOM Component Detail].QTYUsage, Inventory.MinStock, Inventory.Manufacture ORDER BY [BOM Component Detail].NoItem", CNN
With RcBOM.DBRecordset
     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For k = 0 To UBound(Avdata, 2)
            If CBool(Avdata(7, k)) = True Then
               Call ReadDetailBOMBOM(Avdata(0, k), TglRequireDate, Avdata(5, k), PlanningDays)
            Else
               mVarStock = CekStock(Avdata(0, k))
               If (CDbl(Avdata(5, k)) * CDbl(QtyOrder)) > mVarStock Then
                  mVarQTYPO = (CDbl(Avdata(5, k)) * CDbl(QtyOrder)) - mVarStock
                  If mVarQTYPO < 0 Then mVarQTYPO = mVarQTYPO * (-1)
                  If mVarQTYPO >= CDbl(Avdata(4, k)) Then
                     mVarSug = mVarQTYPO + (CDbl(Avdata(5, k)) * CDbl(QtyOrder)) + CDbl(Avdata(3, k))
                  Else
                     mVarSug = mVarQTYPO
                  End If
                  'Dgn Lead
'                    SendDataToServer (" INSERT INTO [Planned Order]" & _
'                                      " (NoItem, [DESC], M_OR_P, PartnerID, [Suggest QTY], [Order QTY], [Required Date], [Suggest Order Date], [Order Date])" & _
'                                      " VALUES   (N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(7, k)) & ", N'" & Avdata(2, k) & "', " & mVarSug & ", " & mVarSug & ", CONVERT(DATETIME, '" & Format(CDate(TglRequireDate), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) + CDbl(Text1)), "dd/mm/yy") & "', 3), CONVERT(DATETIME,'" & Format(CDate(TglRequireDate) + CDbl(Avdata(6, k)), "dd/mm/yy") & "', 3))")
'
'
'               Else
'                    SendDataToServer (" INSERT INTO [Planned Order]" & _
'                                      " (NoItem, [DESC], M_OR_P, PartnerID, [Suggest QTY], [Order QTY], [Required Date], [Suggest Order Date], [Order Date])" & _
'                                      " VALUES   (N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(7, k)) & ", N'" & Avdata(2, k) & "', " & (CDbl(Avdata(5, k)) * CDbl(QtyOrder)) & ", " & (CDbl(Avdata(5, k)) * CDbl(QtyOrder)) & ", CONVERT(DATETIME, '" & Format(CDate(TglRequireDate), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) + CDbl(Text1)), "dd/mm/yy") & "', 3), CONVERT(DATETIME,'" & Format(CDate(TglRequireDate) + CDbl(Avdata(6, k)), "dd/mm/yy") & "', 3))")
                  'Tanpa Lead
                    SendDataToServer (" INSERT INTO [Planned Order]" & _
                                     " (NoItem, [DESC], M_OR_P, PartnerID, [Suggest QTY], [Order QTY], [Required Date], [Suggest Order Date], [Order Date])" & _
                                      " VALUES   (N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(7, k)) & ", N'" & Avdata(2, k) & "', " & mVarSug & ", " & mVarSug & ", CONVERT(DATETIME, '" & Format(CDate(TglRequireDate), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) - CDbl(PlanningDays)), "dd/mm/yy") & "', 3), CONVERT(DATETIME,'" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) - CDbl(PlanningDays)), "dd/mm/yy") & "', 3))")
                  
               
               Else
                    mVarQTYPO = (CDbl(Avdata(5, k)) * CDbl(QtyOrder)) - mVarStock
                    If mVarQTYPO < 0 Then mVarQTYPO = mVarQTYPO * (-1)
                    If mVarQTYPO >= CDbl(Avdata(4, k)) Then
                       mVarSug = mVarQTYPO + (CDbl(Avdata(5, k)) * CDbl(QtyOrder)) + CDbl(Avdata(3, k))
                    Else
                       mVarSug = mVarQTYPO
                    End If
                    SendDataToServer (" INSERT INTO [Planned Order]" & _
                                      " (NoItem, [DESC], M_OR_P, PartnerID, [Suggest QTY], [Order QTY], [Required Date], [Suggest Order Date], [Order Date])" & _
                                      " VALUES   (N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(7, k)) & ", N'" & Avdata(2, k) & "', " & (CDbl(Avdata(5, k)) * CDbl(QtyOrder)) & ", " & (CDbl(Avdata(5, k)) * CDbl(QtyOrder)) & ", CONVERT(DATETIME, '" & Format(CDate(TglRequireDate), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) - CDbl(PlanningDays)), "dd/mm/yy") & "', 3), CONVERT(DATETIME,'" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) - CDbl(PlanningDays)), "dd/mm/yy") & "', 3))")
               
               End If
            End If
        Next
     End If
End With
Set Avdata = Nothing
End Sub

Private Sub ReadDetailBOMBOM(ByVal Param As String, ByVal TglRequireDate As String, ByVal QtyOrder As Long, ByVal PlanningDays As Long)
Dim RcBOM As New DBQuick
Dim Avdata As Variant
Dim mVarStock As Variant
Dim mVarSug As Variant
Dim mParam As String
Dim k As Integer

'RcBOM.DBOpen " SELECT [BOM Component Detail].NoItem, Inventory.ItemName, Inventory.PartnerID, Inventory.ROP AS RQTY, Inventory.MinStock AS [Min Stock],                        [BOM Component Detail].QTYUsage AS [Quote QTY], Inventory.LeadTimeDays, Inventory.Manufacture  FROM         [BOM Component Detail] INNER JOIN                       Inventory ON [BOM Component Detail].NoItem = Inventory.NoItem WHERE     ([BOM Component Detail].Component = N'" & Param & "') GROUP BY Inventory.LeadTimeDays, Inventory.PartnerID, Inventory.ROP, [BOM Component Detail].NoItem, Inventory.ItemName,                        [BOM Component Detail].QTYUsage, Inventory.MinStock, Inventory.Manufacture ORDER BY [BOM Component Detail].NoItem", Cnn
'RcBOM.DBOpen " SELECT [BOM Component Detail].NoItem, Inventory.ItemName, Inventory.PartnerID, Inventory.ROP AS RQTY, Inventory.MinStock AS [Min Stock],                        [BOM Component Detail].QTYUsage AS [Quote QTY], Inventory.LeadTimeDays, Inventory.Manufacture  FROM         [BOM Component Detail] INNER JOIN                       Inventory ON [BOM Component Detail].NoItem = Inventory.NoItem WHERE     ([BOM Component Detail].Component = N'" & Param & "') GROUP BY Inventory.LeadTimeDays, Inventory.PartnerID, Inventory.ROP, [BOM Component Detail].NoItem, Inventory.ItemName,                        [BOM Component Detail].QTYUsage, Inventory.MinStock, Inventory.Manufacture ORDER BY [BOM Component Detail].NoItem", Cnn
RcBOM.DBOpen "SELECT     [BOM Component Detail].Component AS NoItem, Inventory.ItemName, Inventory.PartnerID, Inventory.ROP AS RQTY, Inventory.MinStock AS [Min Stock],                       [BOM Component Detail].QTYUsage AS [Quote QTY], Inventory.LeadTimeDays, Inventory.Manufacture FROM         [BOM Component Detail] INNER JOIN Inventory ON [BOM Component Detail].Component = Inventory.NoItem WHERE     ([BOM Component Detail].NoItem = N'" & Param & "') GROUP BY Inventory.LeadTimeDays, Inventory.PartnerID, Inventory.ROP, Inventory.ItemName, [BOM Component Detail].QTYUsage, Inventory.MinStock,  Inventory.Manufacture, [BOM Component Detail].Component ORDER BY Inventory.Manufacture,[BOM Component Detail].Component", CNN
With RcBOM.DBRecordset
     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For k = 0 To UBound(Avdata, 2)
'            SendDataToServer (" INSERT INTO [Planned Order]" & _
'                              " (NoItem, [DESC], M_OR_P, PartnerID, [Suggest QTY], [Order QTY], [Required Date], [Suggest Order Date], [Order Date])" & _
'                              " VALUES   (N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(7, k)) & ", N'" & Avdata(2, k) & "', 1, 1, CONVERT(DATETIME, '" & Format(CDate(TglRequireDate), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format(CDate(TglRequireDate) - CDbl(Avdata(6, k)), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format(CDate(TglRequireDate) + CDbl(Avdata(6, k)), "dd/mm/yy") & "' , 3))")
'
               mVarStock = CekStock(Avdata(0, k))
               If (CDbl(Avdata(5, k)) * CDbl(QtyOrder)) > mVarStock Then
                  QtyOrder = (CDbl(Avdata(5, k)) * CDbl(QtyOrder)) - mVarStock
                  If QtyOrder < 0 Then QtyOrder = QtyOrder * (-1)
                  If QtyOrder >= CDbl(Avdata(4, k)) Then
                     mVarSug = (CDbl(Avdata(5, k)) * CDbl(QtyOrder)) + CDbl(Avdata(3, k))
                  Else
                     mVarSug = QtyOrder
                  End If
                  'Tambah no Iki Lek Require date Butuh  Lead Time Maneh -> CDbl(Avdata(6, k))
                  SendDataToServer (" INSERT INTO [Planned Order]" & _
                                    " ([Quote Qty],NoItem, [DESC], M_OR_P, PartnerID, [Suggest QTY], [Order QTY], [Required Date], [Suggest Order Date], [Order Date])" & _
                                    " VALUES   (" & (CDbl(Avdata(5, k)) * CDbl(QtyOrder)) & ",N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(7, k)) & ", N'" & Avdata(2, k) & "', " & mVarSug & ", " & mVarSug & ", CONVERT(DATETIME, '" & Format(CDate(TglRequireDate), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) - CDbl(PlanningDays)), "dd/mm/yy") & "', 3), CONVERT(DATETIME,'" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) - CDbl(PlanningDays)), "dd/mm/yy") & "', 3))")
                  
               
               Else
               'Tambah no Iki Lek Require date Butuh  Lead Time Maneh -> CDbl(Avdata(6, k))
                  QtyOrder = (CDbl(Avdata(5, k)) * CDbl(QtyOrder)) - mVarStock
                  If QtyOrder < 0 Then QtyOrder = QtyOrder * (-1)
                  If QtyOrder >= CDbl(Avdata(4, k)) Then
                     mVarSug = (CDbl(Avdata(5, k)) * CDbl(QtyOrder)) + CDbl(Avdata(3, k))
                  Else
                     mVarSug = QtyOrder
                  End If
                  SendDataToServer (" INSERT INTO [Planned Order]" & _
                                    " ([Quote Qty],NoItem, [DESC], M_OR_P, PartnerID, [Suggest QTY], [Order QTY], [Required Date], [Suggest Order Date], [Order Date])" & _
                                    " VALUES   (" & (CDbl(Avdata(5, k)) * CDbl(QtyOrder)) & ",N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(7, k)) & ", N'" & Avdata(2, k) & "', " & CDbl(mVarSug) & ", " & CDbl(mVarSug) & ", CONVERT(DATETIME, '" & Format(CDate(TglRequireDate), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) - CDbl(PlanningDays)), "dd/mm/yy") & "', 3), CONVERT(DATETIME,'" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) - CDbl(PlanningDays)), "dd/mm/yy") & "', 3))")
               End If
        Next
     End If
End With
Set Avdata = Nothing
End Sub

Private Function CekStock(ByVal NoItem As String) As Long
Dim RcCek As New Recordset
RcCek.CursorLocation = adUseClient
'RcCek.Open "SELECT     SUM(QTY_IN) - SUM(QTY_OUT) AS Stock FROM         [Inventory Tabel] WHERE     (NoItem = N'" & NoItem & "')", Cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
RcCek.Open "SELECT     SUM([QTY On Hand]) + SUM([Qty ON Purchase]) + SUM([Qty ON Production]) AS Total FROM         [QTY Availlable] WHERE     (NoItem = N'" & NoItem & "')", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
With RcCek
     If .Recordcount <> 0 Then
        CekStock = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        CekStock = 0
     End If
     .Close
End With
Set RcCek = Nothing
End Function
