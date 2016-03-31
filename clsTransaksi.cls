VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransaksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'Dim Tombolku As TombolTransaksi
'Dim mVarParam As String

'Public Sub OpenParam(ByVal Param As String, ByVal C_OR_S As TombolTransaksi)
'mVarParam = Param
'End Sub

'Public Function PrepareDetail(ByVal ParameterString As String, Optional C_OR_S As TombolTransaksi) As Recordset
'CloseDB PrepareDetail
'Set PrepareDetail = New Recordset
'PrepareDetail.CursorLocation = adUseClient
'If ParameterString = "" Then ParameterString = "xxxxxxxx"
'With PrepareDetail
'     Select Case C_OR_S
'            Case tmbTransaksiDetailPO:
'                 .Open "SELECT [Detail PO].NoItem, Inventory.ItemName, [Detail PO].ItemSupplierID, [Detail PO].QTYPO, [Detail PO].POPrice, [Detail PO].VAT, [Detail PO].ScheduleDate,  [Detail PO].PurchaseID FROM  [Detail PO] INNER JOIN                      Inventory ON [Detail PO].NoItem = Inventory.NoItem WHERE     ([Detail PO].PurchaseID = N'" & ParameterString & "') ORDER BY [Detail PO].NoItem", CNN, adOpenStatic, adLockBatchOptimistic, adCmdText
'            Case tmbTransaksiSC:
'
'     End Select
'     Set .ActiveConnection = Nothing
'End With
'Tombolku = C_OR_S
'End Function

Public Sub PrepareVoucher(ByVal TransID As String, _
                          ByVal PartnerId As String, _
                          ByVal RefNotes As String, _
                          ByVal DateTrans As String, _
                          ByVal Debet As Variant, _
                          ByVal Credit As Variant, _
                          ByVal PurchaseID As String, _
                          ByVal TypeTrans As String)

SendDataToServer (" INSERT INTO [Voucher Batch] (NoIDX, TransID, PartnerID, DateTrans, DateIssued, Debet, Credit, RefNotes, TypeTrans,PurchaseID) " & _
                  " VALUES     (NEWID(), N'" & TransID & "', N'" & PartnerId & "', convert(datetime,'" & Format(DateTrans, "dd/mm/yy") & "',3), convert(datetime,'" & Format(dDateBegin, "dd/mm/yy") & "',3), " & CDbl(Round(Debet, 0)) & ", " & CDbl(Round(Credit, 0)) & ", N'" & Left(RefNotes, 200) & "', N'" & TypeTrans & "',N'" & PurchaseID & "')")
                  
End Sub

Public Sub PreparePayVoucher(ByVal TransID As String, _
                             ByVal Referense As String, _
                             ByVal Debet As Currency, _
                             ByVal Credit As Currency, _
                             ByVal PartnerId As String, _
                             ByVal Noref As String)
Dim mVarSuccess As Boolean
Select Case Left(TransID, 2)
       Case "BP":   'payable
            mVarSuccess = SendDataToServer(" INSERT INTO [Detail TransData] (IDX, TransID,  Debet, Credit, Referense,DNID) " & _
                              " VALUES (NEWID(), N'" & TransID & "', " & CCur(Debet) & ", " & CCur(Credit) & ", N'" & Trim(Left(Referense, 50)) & "','" & Noref & "')")
                              
       Case "BR":   'receivable
            mVarSuccess = SendDataToServer(" INSERT INTO [Detail TransData] (IDX, TransID,  Debet, Credit, Referense,DNID) " & _
                              " VALUES (NEWID(), N'" & TransID & "', " & CDbl(Credit) & ", " & CCur(Debet) & ", N'" & Trim(Left(Referense, 50)) & "','" & Noref & "')")
       Case "BK":
'            mVarSuccess = SendDataToServer(" INSERT INTO [BKK Karyawan]" & _
                                           " ([No Piutang], DateTrans, EmpID, Jumlah, Angsuran, TypeTrans)" & _
                                           " VALUES     (N'1', CONVERT(DATETIME, '" & Format(, "dd/mm/yy") & "', 3), N'1', 1, 1, N'PK')")
                                           
            mVarSuccess = SendDataToServer(" INSERT INTO [Detail TransData] (IDX, TransID,  Debet, Credit, Referense,DNID) " & _
                              " VALUES (NEWID(), N'" & TransID & "', " & CDbl(Credit) & ", " & CCur(Debet) & ", N'" & Trim(Left(Referense, 50)) & "','" & Noref & "')")
       
       Case "BU":
            mVarSuccess = SendDataToServer(" INSERT INTO [Detail TransData] (IDX, TransID,  Debet, Credit, Referense,DNID) " & _
                              " VALUES (NEWID(), N'" & TransID & "', " & CDbl(Credit) & ", " & CCur(Debet) & ", N'" & Trim(Left(Referense, 50)) & "','" & Noref & "')")
       
       
End Select
If mVarSuccess = True Then
   If SendDataToServer(" UPDATE [Voucher Batch] SET Credit = 0 ") = True Then SendDataToServer ("DELETE FROM [Voucher Batch] WHERE (Status = N'1') ")
End If
End Sub

Public Sub PrepareAPItem(ByVal NoItem As String, _
                         ByVal QTY_IN As Double, _
                         ByVal PriceIn As Currency, _
                         ByVal RefTrans As String, _
                         ByVal DateTrans As String, _
                         ByVal TypeTrans As String, _
                         ByVal Discount As Single, _
                         ByVal PPN As Single, _
                         ByVal wHouse As String)
                  
                         
SendDataToServer (" INSERT INTO [Inventory Tabel](NoIdx, NoItem, QTY_IN, PriceIn, QTY_OUT, PriceOut, RefTrans, DateTrans, DateIssued, StockTmp, LockFIFO,TypeTrans,Hpp,Dsc,TX,lokasigdg)" & _
                  " VALUES (NEWID(), N'" & NoItem & "', " & QTY_IN & ", " & PriceIn & ", 0, 0, N'" & RefTrans & "', CONVERT(DATETIME, '" & Format(DateTrans, "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format(dDateBegin, "dd/mm/yy") & "', 3), " & QTY_IN & ", 0,N'" & TypeTrans & "'," & PriceIn & "," & Discount & "," & PPN & ",'" & wHouse & "')")
End Sub

Public Sub PrepareARItem(ByVal NoItem As String, _
                         ByVal QTY_OUT As Variant, _
                         ByVal PriceOut As Currency, _
                         ByVal RefTrans As String, _
                         ByVal DateTrans As String, _
                         ByVal PriceHpp As Variant, _
                         ByVal TypeTrans As String)
                         
SendDataToServer (" INSERT INTO [Inventory Tabel](NoIdx, NoItem, QTY_IN, PriceIn, QTY_OUT, PriceOut, RefTrans, DateTrans, DateIssued, StockTmp, LockFIFO,TypeTrans,Hpp)" & _
                  " VALUES (NEWID(), N'" & NoItem & "',0,0, " & QTY_OUT & ", " & PriceOut & ", N'" & RefTrans & "', CONVERT(DATETIME, '" & Format(DateTrans, "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format(dDateBegin, "dd/mm/yy") & "', 3), " & QTY_OUT & ", 0,N'" & TypeTrans & "'," & PriceHpp & ")")
End Sub

'Public Sub PrepareFifo()
'
'End Sub

Public Function PrepareIndex(ByVal IndexTransaksi As TombolTransaksi, ByVal LenIdx As Integer, Optional ByVal ParamString As String, Optional ByVal IdxString As String)
Dim idx As New DBQuick
Dim mVarNo As Long
Dim StrIdx As String
Dim mThn As String
Select Case IndexTransaksi
       Case tmbTransaksiPO: StrIdx = "SELECT     MAX(RIGHT(PurchaseID, 5)) AS MaxNOm FROM         [PO Order] WHERE     (LEFT(PurchaseID, 2) = N'PO') "
       Case tmbTransaksiReceive: StrIdx = "SELECT MAX(RIGHT(TransID, 5)) AS MaxNOm FROM   TransData GROUP BY TypeTrans HAVING (TypeTrans = N'AP') "
       Case tmbTransaksiSC: StrIdx = "SELECT     MAX(RIGHT(PurchaseID, 5)) AS MaxNOm FROM         [PO Order] WHERE     (LEFT(PurchaseID, 2) = N'SC')"
       Case tmbTransaksiAR: StrIdx = "SELECT MAX(RIGHT(TransID, 5)) AS MaxNOm FROM   TransData GROUP BY TypeTrans HAVING (TypeTrans = N'AR') "
       Case tmbVoucher: StrIdx = "SELECT MAX(RIGHT(TransID, 5)) AS MaxNOm FROM   TransData GROUP BY TypeTrans HAVING (TypeTrans = N'" & ParamString & "') "
       Case tmbDeliveryNotes: StrIdx = "SELECT MAX(RIGHT(TransID, 5)) AS MaxNOm FROM     [TransData] where TypeTrans=N'DN' "
       Case tmbFreight: StrIdx = "SELECT MAX(RIGHT(TypeLoco, 5)) AS MaxNOm FROM [Type Bayar]"
       Case tmbAsmOrder:
            If ParamString <> "" Then
               StrIdx = "SELECT     MAX(SUBSTRING([Prepare Inventory].[Asm Order], " & Len(ParamString) + 2 & ", 15)) AS [Asm Order] FROM  [Prepare Inventory] INNER JOIN [Inventory Group] ON [Prepare Inventory].NoGroup = [Inventory Group].NoGroup WHERE     ([Prepare Inventory].NoGroup = N'" & ParamString & "')"
            Else
               StrIdx = "SELECT MAX([Prepare Inventory].[Asm Order]) AS [Asm Order] FROM [Prepare Inventory] INNER JOIN [Inventory Group] ON [Prepare Inventory].NoGroup = [Inventory Group].NoGroup"
            End If
       Case tmbTransaksiReturBeli: StrIdx = "SELECT     MAX(RIGHT(ReturID, 5)) AS MaxNOm FROM         RETURdATA GROUP BY TypeTrans HAVING      (TypeTrans = N'RB')"
       Case tmbTransaksiReturjual: StrIdx = "SELECT     MAX(RIGHT(ReturID, 5)) AS MaxNOm FROM         RETURdATA GROUP BY TypeTrans HAVING      (TypeTrans = N'RJ')"
       Case tmbTransaksiMutasiPenjualan:
            If ParamString <> "" Then
               StrIdx = "SELECT     MAX(SUBSTRING(NoItem, " & Len(ParamString) + 2 & ", 15)) AS Expr1 FROM         Inventory WHERE     (NoGroup = N'" & ParamString & "') "
            Else
               StrIdx = "SELECT     MAX(NoItem ) AS Expr1 FROM         Inventory WHERE     (NoGroup = N'" & ParamString & "')"
            End If
       Case tmbTransaksiPiutangKaryawan:  StrIdx = "SELECT MAX(RIGHT([No Piutang], 5)) AS MaxNOm FROM [BKK Karyawan] where TypeTrans=N'PPK'"
       Case tmbTransaksiBayarPiutangKaryawan:  StrIdx = "SELECT MAX(RIGHT([No Piutang], 5)) AS MaxNOm FROM [BKK Karyawan] where TypeTrans=N'BPK'"
       Case tmbTransaksiBeliAktivaTetap: StrIdx = "SELECT MAX(RIGHT([No FA], 5)) AS MaxNOm FROM [TR Aktiva Tetap] where typetrans='FB'"
       Case tmbTransaksiJualAktivaTetap: StrIdx = "SELECT MAX(RIGHT([No FA], 5)) AS MaxNOm FROM [TR Aktiva Tetap] where typetrans='FJ'"
       Case tmbTransaksiJournal: StrIdx = "SELECT MAX(RIGHT([JournalID], 5)) AS MaxNOm FROM [Table Journal]"
       Case tmbTransaksiBiayaKeluarJournal: StrIdx = "SELECT MAX(RIGHT([JournalID], 5)) AS MaxNOm FROM  [Table Journal] where typetrans='BKK BIAYA'"
       Case tmbTransaksiBiayaMasukJournal: StrIdx = "SELECT MAX(RIGHT([NoTrans], 5)) AS MaxNOm FROM  [TR Biaya] where typetrans='EM'"
       Case tmbTransaksiBKM: StrIdx = "SELECT MAX(RIGHT([JournalID], 5)) AS MaxNOm FROM  [Table Journal] where typetrans=N'BKM'"
       Case tmbTransaksiBKK: StrIdx = "SELECT MAX(RIGHT([JournalID], 5)) AS MaxNOm FROM  [Table Journal] where typetrans=N'BKK' OR typetrans=N'BKKAT'"
       Case tmbTransaksiBKKKARYAWAN: StrIdx = "SELECT MAX(RIGHT([JournalID], 5)) AS MaxNOm FROM  [Table Journal] where typetrans=N'PPK'"
       Case tmbTransaksiBKMKARYAWAN: StrIdx = "SELECT MAX(RIGHT([JournalID], 5)) AS MaxNOm FROM  [Table Journal] where typetrans=N'BPK'"
       Case tmbTransaksiMutasiGudang: StrIdx = "SELECT MAX(RIGHT([NoMutasi], 5)) AS MaxNOm FROM  [Mutasi Gudang] "
       Case tmbInventory:
            If ParamString <> "" Then
               StrIdx = "SELECT     MAX(SUBSTRING(NoItem, " & Len(ParamString) + 2 & ", 15)) AS Expr1 FROM         Inventory WHERE     (NoGroup = N'" & ParamString & "') "
            Else
               StrIdx = "SELECT     MAX(SUBSTRING(NoItem, " & Len(ParamString) + 2 & ", 15)) AS Expr1 FROM         Inventory  "
            End If
       Case tmbTransaksiMemorial: StrIdx = "SELECT  MAX(RIGHT([JournalID], 5)) AS Expr1 FROM         [Table Journal] WHERE     ([TypeTrans]= N'MEMORIAL') "
       Case tmbTransaksiinvMemorial: StrIdx = "SELECT  MAX(RIGHT([JournalID], 5)) AS Expr1 FROM         [Table Journal] WHERE     ([TypeTrans]= N'INVMEMO') "
       Case tmbTransaksiAkumDepre: StrIdx = "SELECT  MAX(RIGHT([JournalID], 5)) AS Expr1 FROM         [Table Journal] WHERE     ([TypeTrans]= N'AKDEP') "
       Case tmbTransaksiBKMAT: StrIdx = "SELECT  MAX(RIGHT([JournalID], 5)) AS Expr1 FROM         [Table Journal] WHERE     ([TypeTrans]= N'BKMAT') "
       Case tmbTransaksiBKKAT: StrIdx = "SELECT  MAX(RIGHT([JournalID], 5)) AS Expr1 FROM         [Table Journal] WHERE     ([TypeTrans]= N'BKkAT') "
       Case tmbTransaksiInvADJ: StrIdx = "SELECT     MAX(RIGHT(RefTrans, 5)) AS MaxNom FROM         [Inventory Tabel] WHERE     (TypeTrans = N'INVADJ')"
       Case tmbTransaksiInvSUB: StrIdx = "SELECT     MAX(RIGHT(RefTrans, 5)) AS MaxNom FROM         [Inventory Tabel] WHERE     (TypeTrans = N'INVSUB')"
       Case tmbTransaksiNOJOURNAL: StrIdx = "SELECT     MAX(RIGHT(NoUrut, 9)) AS Expr1, YEAR(GETDATE()) AS Tahun FROM         [Table Journal] WHERE     (YEAR(DateIssued) = YEAR(GETDATE()))"
       Case tmbTransaksiChange: StrIdx = "SELECT  MAX(RIGHT([JournalID], 5)) AS Expr1 FROM         [Table Journal] WHERE     ([TypeTrans]= N'CHANGE') "
End Select
'MsgBox StrIdx
idx.DBOpen StrIdx, CNN, lckLockReadOnly
With idx
     If .Recordcount <> 0 Then
        If IsNumeric(IIf(Not IsNull(.Fields(0)), .Fields(0), 0)) Then
            mVarNo = CDbl(IIf(Not IsNull(.Fields(0)), .Fields(0), 0)) + 1
        Else
            mVarNo = Val(IIf(Not IsNull(.Fields(0)), .Fields(0), 0)) + 1
        End If
        mThn = IIf(Not IsNull(.Fields(1)), .Fields(1), Format(Year(dDateBegin), "0###"))
     Else
        mVarNo = 1
     End If
     If IndexTransaksi <> tmbTransaksiNOJOURNAL Then
        PrepareIndex = IdxString & KirimNull(LenIdx - Len(Trim(Str(mVarNo)))) & Trim(Str(mVarNo))
     Else
       'MsgBox Len(PrepareIndex)
        PrepareIndex = "JR" & mThn & "-" & KirimNull(9 - Len(Trim(Str(mVarNo)))) & Trim(Str(mVarNo))
     End If
End With
'idx.CloseDB
End Function

Public Function CiptaKaryaHeaderJournal(ByVal JournalID As String, _
                                         ByVal TransID As String, _
                                         ByVal PurchaseID As String, _
                                         ByVal BankID As String, _
                                         ByVal EmpID As String, _
                                         ByVal PartnerId As String, _
                                         ByVal MataUang As String, _
                                         ByVal DateTrans As String, _
                                         ByVal Periode As Integer, _
                                         ByVal TypeTrans As String) As Boolean
If PurchaseID = "" Or PurchaseID = "xxx" Then PurchaseID = "Null" Else PurchaseID = "N'" & PurchaseID & "'"
If BankID = "" Or BankID = "xxx" Then BankID = "Null" Else BankID = "N'" & BankID & "'"
If EmpID = "" Or EmpID = "xxx" Then EmpID = "Null" Else EmpID = "N'" & EmpID & "'"
If PartnerId = "" Or PartnerId = "xxx" Then PartnerId = "Null" Else PartnerId = "N'" & PartnerId & "'"
If MataUang = "" Or MataUang = "xxx" Then MataUang = "IDR"

CiptaKaryaHeaderJournal = SendDataToServer(" INSERT INTO [Table Journal]" & _
                                           " (JournalID, TransID, PurchaseID, BankID, PartnerID,EmpID, Currency, DateTrans,  Periode, TypeTrans)" & _
                                           " VALUES     (N'" & JournalID & "', N'" & TransID & "', " & PurchaseID & ", " & BankID & "," & PartnerId & ", " & EmpID & ", N'" & MataUang & "', CONVERT(DATETIME, '" & Format(DateTrans, "dd/mm/yy") & "', 3), " & Periode & ", N'" & TypeTrans & "')")
End Function

Public Sub CiptaKaryaDetailJournal(ByVal JournalID As String, _
                                    ByVal NoAccount As String, _
                                    ByVal Doc_Reff As String, _
                                    ByVal Debet As Variant, _
                                    ByVal Credit As Variant)
If Doc_Reff = "" Then Doc_Reff = "Is Null"
SendDataToServer (" INSERT INTO [Detail Journal]" & _
                  " (JournalID, NoAccount, [Doc Reff], Debet, Credit) " & _
                  " VALUES   (N'" & JournalID & "', N'" & NoAccount & "', N'" & Doc_Reff & "', " & CCur(Debet) & ", " & CCur(Credit) & ")")
End Sub

Public Function IdxJournalAuto(ByVal mVarKey As String, ByVal TypeJournal As String) As String
IdxJournalAuto = PrepareIndex(tmbTransaksiJournal, 5, mVarKey, TglIndex(TypeJournal))
End Function

Private Function TglIndex(ByVal TypeJournal As String) As String
Dim TglHari As String
Dim TglBulan As String
Dim TglTahun As String
TglIndex = TypeJournal & "-" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End Function

Public Function UploadQuery(ByVal TypeUploadQuery As String, Optional ByVal Params As String) As String
Select Case UCase(TypeUploadQuery)
       Case "SUPPLIER": UploadQuery = "SELECT PartnerDB.PartnerID AS [Partner ID], PartnerDB.CompanyName AS Perusahaan, PartnerDB.Address AS Alamat, " & _
                                      "PartnerDB.City AS Kota, PartnerDB.PostalCode AS [Kode Pos], PartnerDB.Country AS Negara, PartnerDB.Phone AS Telp, " & _
                                      "TermPayment.[Due Date Calculation] AS termPayment, termPayment.code " & _
                                      " FROM  PartnerDB LEFT OUTER JOIN " & _
                                      " TermPayment ON PartnerDB.Term_code = TermPayment.Code " & _
                                      " WHERE     (PartnerDB.PartnerType = N'SUPPLIER') " & _
                                      " ORDER BY PartnerDB.PartnerID"
                                      
       Case "CUSTOMER": UploadQuery = " SELECT PartnerID AS [Partner ID],CompanyName as Perusahaan, Address AS Alamat, City AS Kota, PostalCode AS [Kode Pos], Country AS Negara, Phone AS Telp FROM PartnerDB WHERE (PartnerType = N'CUSTOMER') ORDER BY PartnerID"
       Case "BANK": UploadQuery = "SELECT     Account AS [No Rekening], [Bank Name] AS [Nama Bank], Currency AS [Mata Uang], [Default] FROM         [Bank Partner] WHERE     (PartnerID = N'" & Params & "') ORDER BY [Default], [Bank Name]"
       Case "BARANG JUAL": UploadQuery = "SELECT NoItem AS [No Barang], ItemName AS [Nama Barang], UOM, PPn FROM   Inventory  ORDER BY NoItem"
       Case "FRANCO BELI": UploadQuery = "SELECT     TypeLoco, TypeFreight FROM         [Type Bayar] WHERE     ([Local] = 0) ORDER BY TypeFreight"
       Case "FRANCO JUAL": UploadQuery = "SELECT     TypeLoco, TypeFreight FROM         [Type Bayar] WHERE     ([Local] = 1) ORDER BY TypeFreight"
       Case "MATA UANG": UploadQuery = "Select * from [Currency Setup]"
End Select
End Function

'Public Function TambahNode(ByVal ActiveNodes As MSComctlLib.Nodes, ByVal ActiveNode As MSComctlLib.Node, Optional ByVal TipicalNode As Byte) As Long
'Dim mTmp As Long
'If TipicalNode = 0 Then
'   Select Case ActiveNodes.Count
'          Case Is = 0:
'               With ActiveNodes.Add(, , 1 & "ANAK", "Tambah Departemen")
'                    .Expanded = True
'                    .Bold = True
'               End With
'               SendDataToServer (" INSERT INTO [Tabel Departemen] " & _
'                                 " ([Kode Dep], [Nama Dep], ReportsTo)" & _
'                                 " VALUES (" & KodeDep & ", N'Tambah Departemen', N'TOP')")
'
'          Case Is >= 1:
'               If ActiveNode.Key = "TOP" Then
'                  With ActiveNodes.Add("TOP", tvwChild, KodeDep & "ANAK", ActiveNode.Text)
'
'                  End With
'               Else
'                  With ActiveNodes.Add(ActiveNode.Key, tvwChild, KodeDep & "ANAK", ActiveNode.Text)
'
'                  End With
'               End If
'    End Select
'ElseIf TipicalNode = 1 Then
'   Select Case ActiveNodes.Count
'          Case Is = 0:
''               With ActiveNodes.Add(, , ParentNode, ParentNode)
''                    .Expanded = True
''                    .Bold = True
''               End With
''              SimpanNode ParentNode, "TOP"
'          Case Is >= 1:
'               ActiveNodes.Item(ActiveNode.Index).Text = ActiveNode.Text
'               'SimpanNode KeyNode, ParentNode
'    End Select
'ElseIf TipicalNode = 2 Then
'End If
'TambahNode = mTmp
''End With
'End Function

'Public Function SimpanNode(ByVal NamaDep As String, ByVal ReportsTo As String, Optional ByVal NoIndex As Long, Optional ByVal Tipical As Byte) As Long
'Dim mTmp As Long
'If Tipical = 0 Then
'   mTmp = KodeDep
'   SendDataToServer (" INSERT INTO [Tabel Departemen] " & _
'                     " ([Kode Dep], [Nama Dep], ReportsTo)" & _
'                     " VALUES     (" & mTmp & ", N'" & NamaDep & "', N'" & ReportTo(ReportsTo) & "')")
'
'ElseIf Tipical = 1 Then
'   SendDataToServer (" Update [Tabel Departemen]" & _
'                     " Set [Nama Dep] = N'" & NamaDep & "', ReportsTo = N'" & ReportsTo & "'" & _
'                     " WHERE  ([Kode Dep] = " & NoIndex & ")")
'End If
'SimpanNode = mTmp
'End Function

'Private Function KodeDep() As Long
'Dim Rc As New DBQuick
'Dim I As Long
'Rc.DBOpen "SELECT     MAX([Kode Dep]) AS Expr1 FROM         [Tabel Departemen]", Cnn
'With Rc
'     If .Recordcount <> 0 Then
'        I = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + 1
'     Else
'        I = 1
'     End If
'     KodeDep = Trim(Str(I))
'End With
'End Function

'Private Function ReportTo(ByVal ParamReportTo As String) As Long
'Dim Rc As New DBQuick
'Dim I As Long
'Rc.DBOpen "SELECT     [Kode Dep] FROM         [Tabel Departemen] WHERE     (ReportsTo = N'" & ParamReportTo & "')", Cnn
'With Rc
'     If .Recordcount <> 0 Then
'        ReportTo = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
'     Else
'        ReportTo = 0
'     End If
'End With
'End Function

'Private Function CariKode(ByVal ParamReportTo As String) As Long
'Dim Rc As New DBQuick
'Dim I As Long
'Rc.DBOpen "SELECT     [Kode Dep] FROM         [Tabel Departemen] WHERE     ([Nama Dep] = N'" & ParamReportTo & "')", Cnn
'With Rc
'     If .Recordcount <> 0 Then
'        CariKode = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
'     Else
'        CariKode = 0
'     End If
'End With
'End Function

Public Function CheckGridKosong(ByVal ActiveRec As Recordset, _
                                Optional ByVal ParamData As String, _
                                Optional ByVal ParamData2 As String, _
                                Optional ByVal TipicalZeroNumber As Boolean = False) As Boolean
Dim RcCek As New DBQuick
Dim i As Integer
Dim j As Integer
Dim Fld As Field
Dim Avdata As Variant
Set RcCek.DBRecordset = ActiveRec.Clone(adLockReadOnly)
With RcCek.DBRecordset
     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For i = 0 To UBound(Avdata, 2)
            j = 0
            For Each Fld In .Fields
               If UCase(ParamData) <> UCase(Fld.Name) And UCase(ParamData2) <> UCase(Fld.Name) Then
                  If Fld.Type <> adBoolean Then
                  If IsNull(Avdata(j, i)) = True Then
                     'MsgBox Fld.Name
                     CheckGridKosong = True
                     GoTo Hell
                  End If
                  End If
               End If
               j = j + 1
            Next
        Next i
     End If
     .Close
End With
Hell:
    Set RcCek = Nothing
    Set Avdata = Nothing
End Function

Public Sub EditHeaderRN(ByVal pPurchaseID As String, _
                        ByVal pEmpID As String, _
                        ByVal pCurrID As String, _
                        ByVal pPartnerId As String, _
                        ByVal pTermPayment As Integer, _
                        ByVal pDiscount As Integer, _
                        ByVal pKurs As Currency, _
                        ByVal pStatus As Boolean, ByVal DetailRecordsetRN As Recordset)
Dim RcCekData As New DBQuick


RcCekData.DBOpen "SELECT     PurchaseID FROM   TransData WHERE     (PurchaseID = N'" & pPurchaseID & "')", CNN, lckLockReadOnly
If RcCekData.DBRecordset.Recordcount <> 0 Then
   SendDataToServer (" UPDATE TransData Set Discount =" & CDbl(pDiscount) & ", EmpID = N'" & pEmpID & "', TermPayment = " & CDbl(pTermPayment) & ", Kurs = " & CDbl(pKurs) & ", CurrID = N'" & pCurrID & "', Status = " & BoolToInt(pStatus) & ", PartnerId = N'" & pPartnerId & "'" & _
                     " WHERE (PurchaseID = N'" & pPurchaseID & "')")
   With DetailRecordsetRN
        If .Recordcount <> 0 Then
           .MoveFirst
           Do
             If .EOF Then Exit Do

                If .status = 2 Then
                   SendDataToServer (" UPDATE [Detail TransData] SET Price = " & CDbl(.Fields("POPrice")) & ",HPP=" & CDbl(.Fields("POPrice")) & ", VAT = " & CDbl(.Fields("VAT")) & " " & _
                                     " WHERE (DNID = N'" & pPurchaseID & "') AND (NoItem = N'" & .Fields("NoItem") & "')")
                End If
                
               .MoveNext
           Loop
'           Set .ActiveConnection = Cnn
'           .Requery
'           Set .ActiveConnection = Nothing
           .MoveLast
        End If
   End With
End If
End Sub


'PENENTUAN PERIODE TRANSAKSI BERJALAN
Public Function GetPeriodeStatus(DateSession As Date) As Boolean
On Error GoTo GetErr
Dim rsRange As New DBQuick
Dim i As Long
Dim strSQL As String

rsRange.DBOpen "SELECT * FROM settingperiod ORDER BY GLFile ASC", CNN, lckLockReadOnly
With rsRange
   If .Recordcount = 0 Then
      GetPeriodeStatus = False
      Exit Function
   End If
   .MoveTopRecord
'   DateSession = Format(DateSession, "dd/mm/yyyy")
   GetPeriodeStatus = False
   Do While Not .DBRecordset.EOF
'        MsgBox DateSession
      If DateSession >= .Fields("BeginDate") And DateSession <= .Fields("EndDate") Then
         GetPeriodeStatus = .Fields("Closed")
         FirstTgl = .Fields("BeginDate")
         LastTgl = .Fields("EndDate")
         Exit Do
      End If
      .MoveNextRecord
   Loop
   Set rsRange = Nothing
   .CloseDB
End With

Exit Function

GetErr:
   MsgBox Err.Description, vbCritical, "GetPeriode Status"
End Function

Public Sub GetTglPeriode()
On Error GoTo GetErr
Dim rsRange As New DBQuick
Dim i As Long

rsRange.DBOpen "SELECT * FROM settingperiod ORDER BY GLFile ASC", CNN, lckLockReadOnly
With rsRange
   If .Recordcount = 0 Then
'      GetPeriodeStatus = False
      Exit Sub
   End If
   .MoveTopRecord
'   DateSession = Format(DateSession, "dd/mm/yyyy")
'   GetPeriodeStatus = False
   Do While Not .DBRecordset.EOF
      If Date >= .Fields("BeginDate") And Date <= .Fields("EndDate") Then
         StartTgl = .Fields("BeginDate")
         EndTgl = .Fields("EndDate")
         Exit Do
      End If
      .MoveNextRecord
   Loop
   Set rsRange = Nothing
   .CloseDB
End With

Exit Sub

GetErr:
   MsgBox Err.Description, vbCritical, "GetPeriode Status"
End Sub
