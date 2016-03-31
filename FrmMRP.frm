VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMRP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MRP Generator"
   ClientHeight    =   3375
   ClientLeft      =   2580
   ClientTop       =   2100
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMRP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "MRP Generation"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3360
      Left            =   0
      ScaleHeight     =   3360
      ScaleWidth      =   6840
      TabIndex        =   7
      Top             =   0
      Width           =   6840
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAAF6F&
         Caption         =   " Proses "
         Height          =   660
         Left            =   135
         TabIndex        =   11
         Top             =   2610
         Visible         =   0   'False
         Width           =   6600
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   330
            Left            =   105
            TabIndex        =   12
            Top             =   240
            Width           =   6390
            _ExtentX        =   11271
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
      End
      Begin VB.CommandButton CmdOK 
         Caption         =   "&Generate"
         Height          =   495
         Index           =   1
         Left            =   5265
         TabIndex        =   4
         Top             =   1245
         Width           =   1440
      End
      Begin VB.CommandButton CmdOK 
         Caption         =   "&Report"
         Height          =   495
         Index           =   2
         Left            =   5265
         TabIndex        =   5
         Top             =   1770
         Width           =   1440
      End
      Begin VB.CommandButton CmdOK 
         Caption         =   "&Close"
         Height          =   495
         Index           =   0
         Left            =   5265
         TabIndex        =   6
         Top             =   720
         Width           =   1440
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAAF6F&
         Caption         =   " Job Status "
         Height          =   1020
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Top             =   1590
         Width           =   5025
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00EAAF6F&
            Caption         =   "Include New Job Requirement"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   345
            TabIndex        =   2
            Top             =   315
            Width           =   4230
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00EAAF6F&
            Caption         =   "Include Quoted Job Requirement"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   345
            TabIndex        =   3
            Top             =   630
            Width           =   4230
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAAF6F&
         Caption         =   " Days "
         Height          =   1020
         Index           =   1
         Left            =   135
         TabIndex        =   8
         Top             =   555
         Width           =   5025
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2355
            TabIndex        =   1
            Text            =   "0"
            Top             =   375
            Width           =   1800
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Order Tolerance"
            Height          =   195
            Index           =   1
            Left            =   165
            TabIndex        =   9
            Top             =   443
            Width           =   1875
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   135
            X2              =   2500
            Y1              =   690
            Y2              =   690
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   2040
         TabIndex        =   0
         Top             =   75
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   71630851
         CurrentDate     =   38533
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Planning Horizon Days"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   13
         Top             =   135
         Width           =   1590
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   135
         X2              =   2340
         Y1              =   390
         Y2              =   390
      End
   End
End
Attribute VB_Name = "FrmMRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mOrderQTY As Long
Dim StrQtyBom As String
Dim mVarBomFirst As String
Dim mVarBom As Boolean

Private Sub cmdOk_Click(Index As Integer)
Dim EmRec As New ClsRecursive
Select Case Index
       Case 0: Unload Me
       Case 1:
         Frame2.Visible = True
            Screen.MousePointer = vbHourglass
            cmdOk(0).Enabled = False
            cmdOk(1).Enabled = False
            cmdOk(2).Enabled = False
            EmRec.ReadComponentOrder DTPicker1.Value, CDbl(Text1), ProgressBar1
            cmdOk(0).Enabled = True
            cmdOk(1).Enabled = True
            cmdOk(2).Enabled = True
            Screen.MousePointer = vbDefault
            MessageBox "Proses generate MRP selesai.", "MRP GENERATION", msgOkOnly, msgInfo
            ProgressBar1.Value = ProgressBar1.Min
         Frame2.Visible = False
       Case 2:
           CallRPTReport "Suggested Order.Rpt", "Select * from [Suggested Order] "
End Select
End Sub

Private Sub Form_Load()
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
DTPicker1.Value = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmMRP = Nothing
End Sub

Private Function RecursiveCallManufacture(ByVal Param As String) As String
Dim Rc As New DBQuick
Dim k As Long
Dim mCal As Boolean
Dim Avdata As Variant
Dim mVarQTYPO As Variant
Dim mVarStock As Variant
Dim mVarSug As Variant
Rc.DBOpen " SELECT [BOM Component Detail].SeqStageID, [BOM Component Detail].Component AS [Kode Barang], Inventory.ItemName AS [Nama Komponen],  [BOM Component Detail].UOM, Inventory.PartnerID AS [Partner ID], PartnerDB.CompanyName AS [Nama Perusahaan], [BOM Component Detail].QTYUsage,Inventory.Manufacture FROM  [BOM Component Detail] INNER JOIN [BOM Stage Detail] ON [BOM Component Detail].SeqStageID = [BOM Stage Detail].SeqStageID AND [BOM Component Detail].NoItem = [BOM Stage Detail].NoItem AND [BOM Component Detail].BomReff = [BOM Stage Detail].BomReff INNER JOIN" & _
          " Inventory INNER JOIN PartnerDB ON Inventory.PartnerID = PartnerDB.PartnerID ON [BOM Component Detail].BomReff = Inventory.BomReff AND  [BOM Component Detail].Component = Inventory.NoItem WHERE     ([BOM Component Detail].NoItem = N'" & Param & "') GROUP BY [BOM Component Detail].SeqStageID, [BOM Component Detail].Component, Inventory.ItemName, [BOM Component Detail].UOM, Inventory.PartnerID,  PartnerDB.CompanyName, [BOM Component Detail].QTYUsage, [BOM Stage Detail].NoLine,Inventory.Manufacture ORDER BY [BOM Stage Detail].NoLine, [BOM Component Detail].SeqStageID", CNN
'messagebox Rc.DBRecordset.Source
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        .MoveFirst
        Do
          If .EOF Then Exit Do
'          mCal = .Fields("MAnufacture")
          If .Fields("Manufacture") = True Then
             CariManufacture Rc.DBRecordset, Rc.DBRecordset.Fields(0)
          End If
'         If mCal = True Then
'            If mVarBom = False Then
'               StrQtyBom = StrQtyBom & " - " & .Fields(0) & " -> " & .Fields("MAnufacture") & vbCrLf
'               mVarBom = True
'               mVarBomFirst = .Fields(0)
'               Exit Do
'            End If
'         End If
          If Not .EOF Then .MoveNext
        Loop
        .MoveFirst
'        If mVarBom = True Then
'           RecursiveCallManufacture mVarBomFirst
'        Else
'           MessageBox StrQtyBom
'           Exit Function
'        End If
'     Else
'        MessageBox StrQtyBom
'        StrQtyBom = ""
'        Exit Function
     End If
End With
End Function

Private Sub ReadComponentOrder()
Dim Rc As New DBQuick
Dim k As Long
Dim Avdata As Variant
Dim mVarQTYPO As Variant
Dim mVarStock As Variant
Dim mVarSug As Variant
Rc.DBOpen " Shape{SELECT [Manufacture Order].NoItem, [Manufacture Order].OrderName, Inventory.Manufacture, [Ord Comp Detail].[Quote Qty] AS [Quote qty], Inventory.LeadTimeDays, Inventory.PartnerID, Inventory.MinStock AS [Min Stock], Inventory.ROP AS RQTY,  [Order Output Detail].StartDate - 1 AS RequireDate, [Manufacture Order].[QTY Order], [Manufacture Order].OrderID, [Manufacture Order].Note, Inventory.BomReff" & _
          "       FROM Inventory INNER JOIN [Ord Comp Detail] INNER JOIN [Order Output Detail] INNER JOIN [Manufacture Order] ON [Order Output Detail].OrderID = [Manufacture Order].OrderID ON [Ord Comp Detail].OrderID = [Order Output Detail].OrderID AND [Ord Comp Detail].StageID = [Order Output Detail].StageID ON Inventory.NoItem = [Manufacture Order].NoItem WHERE  ([Ord Comp Detail].Complete = 0) AND ([Order Output Detail].StartDate <= CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3)) GROUP BY [Manufacture Order].NoItem, [Manufacture Order].OrderName, [Ord Comp Detail].[Quote Qty], Inventory.LeadTimeDays, Inventory.PartnerID,  Inventory.MinStock, Inventory.ROP, [Order Output Detail].StartDate - 1, Inventory.Manufacture, [Manufacture Order].[QTY Order],  [Manufacture Order].OrderID, [Manufacture Order].Note, Inventory.BomReff ORDER BY [Manufacture Order].OrderID, [Manufacture Order].NoItem} As ParentMenu" & _
          " Append({SELECT [BOM Component Detail].SeqStageID, [BOM Component Detail].Component AS [cOMPONENT], Inventory.ItemName AS [Nama Komponen],  [BOM Component Detail].UOM, Inventory.PartnerID AS [Partner ID], PartnerDB.CompanyName AS [Nama Perusahaan], [BOM Component Detail].QTYUsage,Inventory.Manufacture,[BOM Component Detail].NoItem FROM  [BOM Component Detail] INNER JOIN [BOM Stage Detail] ON [BOM Component Detail].SeqStageID = [BOM Stage Detail].SeqStageID AND [BOM Component Detail].NoItem = [BOM Stage Detail].NoItem AND [BOM Component Detail].BomReff = [BOM Stage Detail].BomReff INNER JOIN" & _
          " Inventory INNER JOIN PartnerDB ON Inventory.PartnerID = PartnerDB.PartnerID ON [BOM Component Detail].BomReff = Inventory.BomReff AND  [BOM Component Detail].Component = Inventory.NoItem  GROUP BY [BOM Component Detail].SeqStageID, [BOM Component Detail].Component, Inventory.ItemName, [BOM Component Detail].UOM, Inventory.PartnerID,  PartnerDB.CompanyName, [BOM Component Detail].QTYUsage, [BOM Stage Detail].NoLine,Inventory.Manufacture,[BOM Component Detail].NoItem ORDER BY [BOM Stage Detail].NoLine, [BOM Component Detail].SeqStageID} Relate NoItem to NoItem)", CNN
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        ProgressBar1.Min = 0
        ProgressBar1.Max = .Recordcount
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        SendDataToServer ("DELETE FROM [Planned Order]")
        For k = 0 To UBound(Avdata, 2)
'            If CBool(Avdata(2, k)) = True Then RecursiveCallManufacture Avdata(0, k)
             ProgressBar1.Value = k + 1
             mOrderQTY = Avdata(9, k)
             mVarStock = CekStock(Avdata(0, k))
             If CDbl(Avdata(3, k)) > mVarStock Then
                mVarQTYPO = CDbl(Avdata(3, k)) - mVarStock
                If mVarQTYPO < 0 Then mVarQTYPO = mVarQTYPO * (-1)
                If mVarQTYPO >= CDbl(Avdata(6, k)) Then
                   mVarSug = mVarQTYPO + CDbl(Avdata(7, k))
'                  mVarSug = mVarQTYPO + CDbl(Avdata(3, k)) + CDbl(Avdata(7, k))
                Else
                   mVarSug = mVarQTYPO
                End If
                  'Tgl Required + lead
'                  SendDataToServer (" INSERT INTO [Planned Order]" & _
'                                    " (OrderID,Note,NoItem, [DESC], M_OR_P, PartnerID, [Suggest QTY], [Order QTY], [Required Date], [Suggest Order Date], [Order Date])" & _
'                                    " VALUES   (N'" & Avdata(10, k) & "',N'" & Avdata(11, k) & "',N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(2, k)) & ", N'" & Avdata(5, k) & "', " & CDbl(mVarSug) & ", " & CDbl(mVarSug) & ", CONVERT(DATETIME, '" & Format(Avdata(8, k), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format((CDbl(Avdata(8, k)) - CDbl(Avdata(4, k)) + CDbl(Text1)), "dd/mm/yy") & "', 3), CONVERT(DATETIME,'" & Format(Avdata(8, k) + CDbl(Avdata(4, k)), "dd/mm/yy") & "', 3))")
'               Else
'                  SendDataToServer (" INSERT INTO [Planned Order]" & _
'                                    " (OrderID,Note,NoItem, [DESC], M_OR_P, PartnerID, [Suggest QTY], [Order QTY], [Required Date], [Suggest Order Date], [Order Date])" & _
'                                    " VALUES   (N'" & Avdata(10, k) & "',N'" & Avdata(11, k) & "',N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(2, k)) & ", N'" & Avdata(5, k) & "', " & CDbl(Avdata(3, k)) & ", 0, CONVERT(DATETIME, '" & Format(Avdata(8, k), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format((CDbl(Avdata(8, k)) - CDbl(Avdata(4, k)) + CDbl(Text1)), "dd/mm/yy") & "', 3), CONVERT(DATETIME,'" & Format(CDbl(Avdata(8, k)) + CDbl(Avdata(4, k)), "dd/mm/yy") & "', 3))")

                  SendDataToServer (" INSERT INTO [Planned Order]" & _
                                    " (OrderID,Note,NoItem, [DESC], M_OR_P, PartnerID,[Quote Qty], [Suggest QTY], [Order QTY], [Required Date], [Suggest Order Date], [Order Date])" & _
                                    " VALUES   (N'" & Avdata(10, k) & "',N'" & Avdata(11, k) & "',N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(2, k)) & ", N'" & Avdata(5, k) & "',, " & CDbl(mVarSug) & ", " & CDbl(mVarSug) & ", CONVERT(DATETIME, '" & Format(Avdata(8, k), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format((CDbl(Avdata(8, k)) - CDbl(Avdata(4, k)) + CDbl(Text1)), "dd/mm/yy") & "', 3), CONVERT(DATETIME,'" & Format((CDbl(Avdata(8, k)) - CDbl(Avdata(4, k)) + CDbl(Text1)), "dd/mm/yy") & "', 3))")
               Else
                  SendDataToServer (" INSERT INTO [Planned Order]" & _
                                    " (OrderID,Note,NoItem, [DESC], M_OR_P, PartnerID, [Suggest QTY], [Order QTY], [Required Date], [Suggest Order Date], [Order Date])" & _
                                    " VALUES   (N'" & Avdata(10, k) & "',N'" & Avdata(11, k) & "',N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(2, k)) & ", N'" & Avdata(5, k) & "', " & CDbl(Avdata(3, k)) & ", 0, CONVERT(DATETIME, '" & Format(Avdata(8, k), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format((CDbl(Avdata(8, k)) - CDbl(Avdata(4, k)) + CDbl(Text1)), "dd/mm/yy") & "', 3), CONVERT(DATETIME,'" & Format((CDbl(Avdata(8, k)) - CDbl(Avdata(4, k)) + CDbl(Text1)), "dd/mm/yy") & "', 3))")

               End If
            If CBool(Avdata(2, k)) = False Then
               ReadBOMBOM Avdata(0, k), Avdata(8, k), Avdata(9, k)
            End If
        Next
        
     End If
End With
Set Avdata = Nothing
End Sub

Private Sub ReadBOMBOM(ByVal Param As String, ByVal TglRequireDate As String, ByVal QtyOrder As Long)
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
               Call ReadDetailBOMBOM(Avdata(0, k), TglRequireDate, Avdata(5, k))
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
                                      " VALUES   (N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(7, k)) & ", N'" & Avdata(2, k) & "', " & mVarSug & ", " & mVarSug & ", CONVERT(DATETIME, '" & Format(CDate(TglRequireDate), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) + CDbl(Text1)), "dd/mm/yy") & "', 3), CONVERT(DATETIME,'" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) + CDbl(Text1)), "dd/mm/yy") & "', 3))")
                  
               
               Else
                    SendDataToServer (" INSERT INTO [Planned Order]" & _
                                      " (NoItem, [DESC], M_OR_P, PartnerID, [Suggest QTY], [Order QTY], [Required Date], [Suggest Order Date], [Order Date])" & _
                                      " VALUES   (N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(7, k)) & ", N'" & Avdata(2, k) & "', " & (CDbl(Avdata(5, k)) * CDbl(QtyOrder)) & ", " & (CDbl(Avdata(5, k)) * CDbl(QtyOrder)) & ", CONVERT(DATETIME, '" & Format(CDate(TglRequireDate), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) + CDbl(Text1)), "dd/mm/yy") & "', 3), CONVERT(DATETIME,'" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) + CDbl(Text1)), "dd/mm/yy") & "', 3))")
               
               End If
            End If
        Next
     End If
End With
Set Avdata = Nothing
End Sub

Private Sub ReadDetailBOMBOM(ByVal Param As String, ByVal TglRequireDate As String, ByVal QtyOrder As Long)
Dim RcBOM As New DBQuick
Dim Avdata As Variant
Dim mVarStock As Variant
Dim mVarSug As Variant
Dim mParam As String
Dim k As Integer
RcBOM.DBOpen " SELECT [BOM Component Detail].NoItem, Inventory.ItemName, Inventory.PartnerID, Inventory.ROP AS RQTY, Inventory.MinStock AS [Min Stock],                        [BOM Component Detail].QTYUsage AS [Quote QTY], Inventory.LeadTimeDays, Inventory.Manufacture  FROM         [BOM Component Detail] INNER JOIN                       Inventory ON [BOM Component Detail].NoItem = Inventory.NoItem WHERE     ([BOM Component Detail].Component = N'" & Param & "') GROUP BY Inventory.LeadTimeDays, Inventory.PartnerID, Inventory.ROP, [BOM Component Detail].NoItem, Inventory.ItemName,                        [BOM Component Detail].QTYUsage, Inventory.MinStock, Inventory.Manufacture ORDER BY [BOM Component Detail].NoItem", CNN
With RcBOM.DBRecordset
     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For k = 0 To UBound(Avdata, 2)
'            SendDataToServer (" INSERT INTO [Planned Order]" & _
'                              " (NoItem, [DESC], M_OR_P, PartnerID, [Suggest QTY], [Order QTY], [Required Date], [Suggest Order Date], [Order Date])" & _
'                              " VALUES   (N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(7, k)) & ", N'" & Avdata(2, k) & "', 1, 1, CONVERT(DATETIME, '" & Format(CDate(TglRequireDate), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format(CDate(TglRequireDate) - CDbl(Avdata(6, k)), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format(CDate(TglRequireDate) + CDbl(Avdata(6, k)), "dd/mm/yy") & "' , 3))")
'
               mVarStock = CekStock(Avdata(0, k))
               If (CDbl(Avdata(5, k)) * CDbl(mOrderQTY)) > mVarStock Then
                  QtyOrder = (CDbl(Avdata(5, k)) * CDbl(mOrderQTY)) - mVarStock
                  If QtyOrder < 0 Then QtyOrder = QtyOrder * (-1)
                  If QtyOrder >= CDbl(Avdata(4, k)) Then
                     mVarSug = QtyOrder + (CDbl(Avdata(5, k)) * CDbl(mOrderQTY)) + CDbl(Avdata(3, k))
                  Else
                     mVarSug = QtyOrder
                  End If
                  'Tambah no Iki Lek Require date Butuh  Lead Time Maneh -> CDbl(Avdata(6, k))
                  SendDataToServer (" INSERT INTO [Planned Order]" & _
                                    " (NoItem, [DESC], M_OR_P, PartnerID, [Suggest QTY], [Order QTY], [Required Date], [Suggest Order Date], [Order Date])" & _
                                    " VALUES   (N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(7, k)) & ", N'" & Avdata(2, k) & "', " & mVarSug & ", " & mVarSug & ", CONVERT(DATETIME, '" & Format(CDate(TglRequireDate), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) + CDbl(Text1)), "dd/mm/yy") & "', 3), CONVERT(DATETIME,'" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) + CDbl(Text1)), "dd/mm/yy") & "', 3))")
                  
               
               Else
               'Tambah no Iki Lek Require date Butuh  Lead Time Maneh -> CDbl(Avdata(6, k))
               SendDataToServer (" INSERT INTO [Planned Order]" & _
                                 " (NoItem, [DESC], M_OR_P, PartnerID, [Suggest QTY], [Order QTY], [Required Date], [Suggest Order Date], [Order Date])" & _
                                 " VALUES   (N'" & Avdata(0, k) & "', N'" & Avdata(1, k) & "', " & BoolToInt(Avdata(7, k)) & ", N'" & Avdata(2, k) & "', " & (CDbl(Avdata(5, k)) * CDbl(mOrderQTY)) & ", " & (CDbl(Avdata(5, k)) * CDbl(mOrderQTY)) & ", CONVERT(DATETIME, '" & Format(CDate(TglRequireDate), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) + CDbl(Text1)), "dd/mm/yy") & "', 3), CONVERT(DATETIME,'" & Format((CDate(TglRequireDate) - CDbl(Avdata(6, k)) + CDbl(Text1)), "dd/mm/yy") & "', 3))")
               
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

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
ValidNum KeyAscii
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Text1 = "" Then Text1 = 0
End Sub

Private Function CariManufacture(ByVal Rec As Recordset, ByVal No_Item As String, Optional MoveRec As Long)
If Rec.Recordcount <> 0 Then
   If MoveRec = 0 Then MoveRec = Rec.AbsolutePosition
   Rec.AbsolutePosition = MoveRec
   If Not Rec.EOF Then
      MoveRec = MoveRec + 1
      CariManufacture = CariManufacture(Rec, Rec.Fields(0), MoveRec)
   End If
End If
End Function
