VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------
'untuk manambahkan item baru tambahkan pada procedure/function :
'  1.TransactionName    -> untuk menambahkan nama transaksi yg baru
'  2.TransactionID      -> untuk menambahkan ID transaksi yg Baru
'  3.SetMainTable       -> untuk menentukan nama table & field untuk ID
'  4.Tambah nilai pada function TransactionCount

'--------------------------------first version-------------------------------
'Public Sub CurrentDepartment(currDept As String)
'Public Function ItemTypeCount() As Integer
'Public Function ItemTypeCount()
'Public Function GetItemType(Index As Integer) As String
'Public Function TransactionCount() As Integer
'Public Function TransactionName(Index As Integer) As Stringd
'Public Function TransactionID(Index As Integer) As String
'Private Sub SetMainTable(IDType As String)
'Public Function GetID(IDType As String) As String
'Function GetDept() As String
'Private Function GetNewID()
'Public Function IsValidID() As Boolean

'-------------------------------21 Maret 2008--------------------------------
'Property ExtParameter
'Tambah Trasaction ID untuk Inventory

'-------------------------------29 Maret 2008--------------------------------
'Tambah TRansaction ID untuk Dryer

'-------------------------------03 April 2008--------------------------------
'Tambah TRansaction ID untuk Lot No

'------------------------------25 April 2008 ----------------------------------
'Tambah Doc No ACID Treatment
'               Alkali Treatment
'               Ekstraksi Auticlave
'               Ekstraksi REaktor
'               Filter Press
'               Gellification
'               Pembungkusan
'-----------------------------------------------------------------------------


Option Explicit
Private rsID As DBQuick
Private rsFormat As DBQuick
Private RsLookup As DBQuick
Private MainTable As String
Private MainField As String
Private aID As String
Private CounterIsExsist As Boolean
Private lNewID As String
Private lCurrentDept As String
Private lExtParameter As String
Private PrefixCounter As String
                                

Public Property Let ExtParameter(Params As String)
    lExtParameter = Params
End Property


Public Sub CurrentDepartment(currDept As String)
   lCurrentDept = currDept
End Sub

'Return the number of Item Type
Public Function ItemTypeCount() As Integer  'counter start from 1 not 0
   ItemTypeCount = 8 '13
End Function



'Return the name of item type refer to index parameter
Public Function GetItemType(Index As Integer) As String
   Select Case Index
      Case 1: GetItemType = "Fix Character"
      Case 2: GetItemType = "Tanggal"
      Case 3: GetItemType = "Bulan"
      Case 4: GetItemType = "Tahun"
      Case 5: GetItemType = "Counter"
      Case 6: GetItemType = "Kode Dept"
      Case 7: GetItemType = "Ext Parameter"
      Case 8: GetItemType = "Bulan Romawi"
   End Select
End Function

'return the number of Transaction on Application
Public Function TransactionCount() As Integer
   TransactionCount = 47
End Function


'return the name of Transaction base on Index Parameter
Public Function TransactionName(Index As Integer) As String
   Select Case Index
        Case 1: TransactionName = "Purchase Order"
        Case 2: TransactionName = "Sales Order"
        Case 3: TransactionName = "Sales Quote"
        Case 4: TransactionName = "Receive Note"
        Case 5: TransactionName = "Factory Request"
        Case 6: TransactionName = "Sales Return"
        Case 7: TransactionName = "Purchase Return"
        Case 8: TransactionName = "Delivery Order"
        Case 9: TransactionName = "Permintaan Penawaran Harga"
        Case 10: TransactionName = "Permintaan Pembelian"
        Case 11: TransactionName = "Tanda Terima Rumput Laut"
        Case 13: TransactionName = "Delevery Order from WareHouse"
        Case 14: TransactionName = "Permintaan Barang"
        Case 15: TransactionName = "Penerimaan Produk Jadi"
        Case 16: TransactionName = "Sales Contract"
        Case 17: TransactionName = "Hydraulic Press"
        Case 18: TransactionName = "Inventory"
        Case 19: TransactionName = "PreLot Number"
        Case 20: TransactionName = "Dryer"
        Case 21: TransactionName = "Serah Terima Produk Jadi"
        Case 22: TransactionName = "ACID Treatment"
        Case 23: TransactionName = "Alkali Treatment"
        Case 24: TransactionName = "Ekstraksi di Autoclave"
        Case 25: TransactionName = "Ekstraksi di Reaktor"
        Case 26: TransactionName = "Filter Press"
        Case 27: TransactionName = "Gellification"
        Case 28: TransactionName = "Pembungkusan"
        Case 29: TransactionName = "Bleaching Treatment"
        Case 30: TransactionName = "Nomor Ekstraksi"
        Case 31: TransactionName = "Rekomendasi Proses Ekstraksi"
        Case 32: TransactionName = "Rencana Pembelian Bahan"
        Case 33: TransactionName = "Permintaan Sample"
        Case 34: TransactionName = "Invoice Penjualan"
        Case 12: TransactionName = "Invoice Pembelian"
        Case 35: TransactionName = "Voucher Penjualan"
        Case 36: TransactionName = "Voucher Pembelian"
        Case 37: TransactionName = "Penerimaan Kas"
        Case 38: TransactionName = "Pengeluaran Kas"
        Case 39: TransactionName = "Jurnal Umum"
        Case 40: TransactionName = "Debit-Credit Note"
        Case 41: TransactionName = "PO Petty Cash"
        Case 42: TransactionName = "Transfer Kas"
        Case 43: TransactionName = "Pengiriman RL"
        Case 44: TransactionName = "RL Batch"
        Case 45: TransactionName = "Manufacture Order"
        Case 46: TransactionName = "Sample Rumput Laut"
        Case 47: TransactionName = "Pemeriksaan Bahan Penunjang"
   End Select
End Function

'return the ID of each transaction base on Index parameter
Public Function TransactionID(Index As Integer) As String
    Select Case Index
        Case 1: TransactionID = "PO"      'Purchase Order
        Case 2: TransactionID = "SO"      'Sales Order
        Case 3: TransactionID = "SQ"      'Sales Quote
        Case 4: TransactionID = "RN"      'Receive Note
        Case 5: TransactionID = "FR"      'Factory Request
        Case 6: TransactionID = "SR"      'Sales Return / retur Customer
        Case 7: TransactionID = "PR"      'Purchase Return / retur Supplier
        Case 8: TransactionID = "DO"      'Delivery Order
        Case 9: TransactionID = "OF"      'Permintaan Penawaran Harga
        Case 10: TransactionID = "PP"     'Permintaan Pembelian
        Case 11: TransactionID = "RR"     'Tanda Terima Rumput Laut
        Case 12: TransactionID = "IS"     'supplier invoice
        Case 13: TransactionID = "DW"     'delevery order warehouse
        Case 14: TransactionID = "MR"     'material request
        Case 15: TransactionID = "TP"     'Penerimaan Produk Jadi
        Case 16: TransactionID = "SC"     'Sales Contract
        Case 17: TransactionID = "HP"     'Hydraulic Press
        Case 18: TransactionID = "IN"     'Inventory
        Case 19: TransactionID = "PL"     'Prelot Number
        Case 20: TransactionID = "DR"     'dryer
        Case 21: TransactionID = "FG"       'Serah terima produk jadi
        Case 22: TransactionID = "ACID"     '"ACID Treatment"
        Case 23: TransactionID = "ALKALI"   '"Alkali Treatment"
        Case 24: TransactionID = "EA"       '"Ekstraksi di Autoclave"
        Case 25: TransactionID = "ER"       '"Ekstraksi di Reaktor"
        Case 26: TransactionID = "FILTER"   '"Filter Press"
        Case 27: TransactionID = "GELL"     '"Gellification"
        Case 28: TransactionID = "PB"       '"Pembungkusan"
        Case 29: TransactionID = "BLEACH"   '"bleaching treatment
        Case 30: TransactionID = "EKS"
        Case 31: TransactionID = "RK"       'Rekomendasi Proses Ekstraksi
        Case 32: TransactionID = "RPB"    'rencana pembelian bahan
        Case 33: TransactionID = "SAMPLE" 'Permintaan Sample
        Case 34: TransactionID = "IP"     ' invoice Penjualan
        Case 35: TransactionID = "VS"     ' voucher penjualan
        Case 36: TransactionID = "VP"     ' voucher pembelian
        Case 37: TransactionID = "KM"     'penerimaan kas
        Case 38: TransactionID = "KK"     'pengeluaran kas
        Case 39: TransactionID = "MM"     'Jurnal umum
        Case 40: TransactionID = "DCN"    'debit-credit note
        Case 41: TransactionID = "PC"     'petty cash
        Case 42: TransactionID = "TK"     'TRANSFER KAS
        Case 43: TransactionID = "SS"     'PENGIRIMAN RL
        Case 44: TransactionID = "RLB"    ' RL BAtch
        Case 45: TransactionID = "MO"    ' Manufacture Order
        Case 46: TransactionID = "SRL"   ' Sample Rumput Laut
        Case 47: TransactionID = "SPJ"   'Pemeriksaan Bahan penunjang
    End Select
End Function

'initiate the table name and field name for ID
Private Sub SetMainTable(IDType As String)

    Select Case IDType
        Case "RPB"
            MainTable = "PO Order"
            MainField = "PurchaseID"
        
        Case "EKS"
            MainTable = "Manufacture Order"
            MainField = "ekstraksi_no"
         
        Case "RK"
            MainTable = "labrekomEkstraksi"
            MainField = "SplNo"
         
        Case "PO"
            MainTable = "PO Order"
            MainField = "PurchaseID"

        Case "SO"
            MainTable = "PO Order"
            MainField = "PurchaseID"

        Case "SQ"
            MainTable = "PO Order"
            MainField = "PurchaseID"

        Case "SC"
            MainTable = "PO Order"
            MainField = "PurchaseID"

        Case "RN"
            MainTable = "TransData"
            MainField = "TransID"

        Case "FR"
            MainTable = "Backflush_header"
            MainField = "IDTrans"

        Case "SR"
            MainTable = "ReturData"
            MainField = "ReturID"

        Case "PR"
            MainTable = "ReturData"
            MainField = "ReturID"

        Case "DO"
            MainTable = "TransData"
            MainField = "TransID"

        Case "OF"
            MainTable = "SPPH_header"
            MainField = "SPPHID"

        Case "SPL"
            MainTable = "LabSample_Header"
            MainField = "SplNo"

        Case "PP"
            MainTable = "SPP_Header"
            MainField = "SPPID"

        Case "RR"
            MainTable = "TransData"
            MainField = "TransID"

        Case "IS"
            MainTable = "invoicing_header"
            MainField = "invoice_no"

        Case "DW"
            MainTable = "TransData"
            MainField = "TransID"

        Case "MR"
            MainTable = "backflush_Header"
            MainField = "IDTrans"

        Case "TP"
            MainTable = "item_finishGood"
            MainField = "no_penerimaan"

        Case "HP"
            MainTable = "PRESS_HEADER"
            MainField = "No_press"

        Case "IN"
            MainTable = "Inventory"
            MainField = "noItem"

        Case "PL"
            MainTable = "mixing_header"
            MainField = "PreLot"

        Case "DR"
            MainTable = "dryer_header"
            MainField = "ID"

        Case "FG"
            MainTable = "backflush_header"
            MainField = "IDTrans"

        Case "ACID"
            MainTable = "labprosesproduksi_header"
            MainField = "dokno"

        Case "ALKALI"
            MainTable = "labprosesproduksi_header"
            MainField = "dokno"

        Case "EA"
            MainTable = "labprosesproduksi_header"
            MainField = "dokno"

        Case "ER"
            MainTable = "labprosesproduksi_header"
            MainField = "dokno"

        Case "FILTER"
            MainTable = "prodprosesproduksi_header"
            MainField = "dokno"

        Case "GELL"
            MainTable = "prodprosesproduksi_header"
            MainField = "dokno"

        Case "PB"
            MainTable = "produksi_pembungkusan"
            MainField = "dokno"

        Case "BLEACH"
            MainTable = "prodprosesproduksi_header"
            MainField = "dokno"
        
        Case "SAMPLE"
            MainTable = "permintaansample"
            MainField = "nomor"
            
        Case "IP", "IS", "VS", "VP":
            MainTable = "TransData"
            MainField = "TransID"
            
        Case "KK", "KM", "JU", "DCN", "TK", "MM":
            MainTable = "Table Journal"
            MainField = "JournalID"
            
        Case "PC", "SS":
            MainTable = "TransData"
            MainField = "TransID"
        
        Case "RLB":
            MainTable = "Inventory Tabel"
            MainField = "sl_no"
'        Case "TK":
'            MainTable = "TransData"
'            MainField = "TransID"
        Case "MO":
            MainTable = "Manufacture Order"
            MainField = "OrderID"
            
        Case "SRL":
            MainTable = "LabSample_rl"
            MainField = "SplNo"
            
       Case "SPJ":
            MainTable = "LabSampleBahan"
            MainField = "ID"
    End Select

End Sub

Public Function GetID(IDType As String) As String
   Dim lDate As String, lMonth As String, lYear As String
   Dim x As Integer
   CounterIsExsist = False
   SetMainTable IDType
   Set rsFormat = New DBQuick
   rsFormat.DBOpen "select * from [Trans Setup] where IDTrans='" & IDType & "' order by [no Index]", CNN
   If rsFormat.Recordcount > 0 Then
      aID = ""
      rsFormat.DBRecordset.MoveFirst
      While Not rsFormat.DBRecordset.EOF
         Select Case rsFormat.DBRecordset.Fields("type").Value
            Case "Fix Character":
               aID = aID & Mid(rsFormat.DBRecordset.Fields("note").Value, _
                             1, rsFormat.DBRecordset.Fields("Length per Account").Value) & _
                     rsFormat.DBRecordset.Fields("Prefix").Value
            Case "Tanggal":
               lDate = Format(Now, "dd")
               If rsFormat.DBRecordset.Fields("Length per Account").Value > 2 Then
                  For x = 1 To rsFormat.DBRecordset.Fields("Length per Account").Value - 2
                     lDate = "0" & lDate
                  Next
               End If
               aID = aID & lDate & rsFormat.DBRecordset.Fields("Prefix").Value
            Case "Bulan":
               lMonth = Format(Now, "MM")
               If rsFormat.DBRecordset.Fields("Length per Account").Value > 2 Then
                  For x = 1 To rsFormat.DBRecordset.Fields("Length per Account").Value - 2
                     lMonth = "0" & lMonth
                  Next
               End If
               aID = aID & lMonth & rsFormat.DBRecordset.Fields("Prefix").Value
            Case "Tahun":
               lYear = Format(Now, "yyyy")
               If rsFormat.DBRecordset.Fields("Length per Account").Value <= 4 Then
                  lYear = Right(lYear, rsFormat.DBRecordset.Fields("Length per Account").Value)
               Else
                  For x = 1 To rsFormat.DBRecordset.Fields("Length per Account").Value - 4
                     lYear = "0" & lYear
                  Next
               End If
               aID = aID & lYear & rsFormat.DBRecordset.Fields("Prefix").Value
            Case "Counter":
               For x = 1 To rsFormat.DBRecordset.Fields("Length per Account").Value
                  aID = aID & "%"
               Next
               'aID = aID
               PrefixCounter = IIf(IsNull(rsFormat.DBRecordset.Fields("Prefix").Value), "", rsFormat.DBRecordset.Fields("Prefix").Value)
               CounterIsExsist = True
            Case "Kode Dept":
               If Len(CurrentDept) > rsFormat.DBRecordset.Fields("Length per Account").Value Then
                  aID = aID & Left(CurrentDept, rsFormat.DBRecordset.Fields("Length per Account").Value) & rsFormat.DBRecordset.Fields("Prefix").Value
               Else
                  aID = aID & CurrentDept & rsFormat.DBRecordset.Fields("Prefix").Value
               End If
            Case "Ext Parameter":
               If Len(lExtParameter) > rsFormat.DBRecordset.Fields("Length per Account").Value Then
                  aID = aID & Left(lExtParameter, rsFormat.DBRecordset.Fields("Length per Account").Value) & rsFormat.DBRecordset.Fields("Prefix").Value
               Else
                  aID = aID & lExtParameter & rsFormat.DBRecordset.Fields("Prefix").Value
               End If
         
            Case "Bulan Romawi":
               Select Case Format(Now, "MM")
                  Case "01": lMonth = "I"
                  Case "02": lMonth = "II"
                  Case "03": lMonth = "III"
                  Case "04": lMonth = "IV"
                  Case "05": lMonth = "V"
                  Case "06": lMonth = "VI"
                  Case "07": lMonth = "VII"
                  Case "08": lMonth = "VIII"
                  Case "09": lMonth = "IX"
                  Case "10": lMonth = "X"
                  Case "11": lMonth = "XI"
                  Case "12": lMonth = "XII"
               End Select
               aID = aID & lMonth & rsFormat.DBRecordset.Fields("Prefix").Value
'
'            Case "Thickness":
'               aID = aID & Mid(rsFormat.DBRecordset.Fields("note").value, _
'                             1, rsFormat.DBRecordset.Fields("Length per Account").value) & _
'                     rsFormat.DBRecordset.Fields("Prefix").value
'
'            Case "Film Type":
'               aID = aID & Mid(rsFormat.DBRecordset.Fields("note").value, _
'                             1, rsFormat.DBRecordset.Fields("Length per Account").value) & _
'                     rsFormat.DBRecordset.Fields("Prefix").value
'
'            Case "Grade SR":
'               aID = aID & Mid(rsFormat.DBRecordset.Fields("note").value, _
'                             1, rsFormat.DBRecordset.Fields("Length per Account").value) & _
'                     rsFormat.DBRecordset.Fields("Prefix").value
'
'            Case "Posisi GR":
'               aID = aID & Mid(rsFormat.DBRecordset.Fields("note").value, _
'                             1, rsFormat.DBRecordset.Fields("Length per Account").value) & _
'                     rsFormat.DBRecordset.Fields("Prefix").value
'
'            Case "Posisi Roll":
'               aID = aID & Mid(rsFormat.DBRecordset.Fields("note").value, _
'                             1, rsFormat.DBRecordset.Fields("Length per Account").value) & _
'                     rsFormat.DBRecordset.Fields("Prefix").value
         
         End Select
         rsFormat.DBRecordset.MoveNext
      Wend
      lNewID = GetNewID
      GetID = lNewID
   Else
      lNewID = " "
      GetID = " "
   End If
End Function


Function GetDept() As String
   Dim rsDept As New DBQuick
   Dim strUser As String
   Dim strEmpID As String
   strUser = MainMenu.StatusBar1.Panels(1).Text
   rsDept.DBOpen "select empID,dept from [user table] where [user name]='" & strUser & "'", CNN, lckLockReadOnly
   If rsDept.DBRecordset.Recordcount > 0 Then
      strEmpID = rsDept.DBRecordset.Fields(0)
       GetDept = rsDept.DBRecordset.Fields(1)
   Else
      GetDept = ""
   End If
   Set rsDept = Nothing
End Function


Private Function GetNewID()
   Dim strSQL As String
   Dim strID As String
   Dim rsNewID As New DBQuick
   Dim startPos As Integer
   Dim lengthNo As Integer
   Dim x As Double
   Dim Y, z As Integer
   Dim newStr As String
   If CounterIsExsist Then
      startPos = 1
      While Not Mid(aID, startPos, 1) = "%"
         startPos = startPos + 1
      Wend
      x = startPos
      lengthNo = 0
      While Mid(aID, x, 1) = "%"
         x = x + 1
         lengthNo = lengthNo + 1
      Wend
      strSQL = "select max(cast(substring([" & MainField & "]," & startPos & "," & lengthNo & ") as integer)) from [" & MainTable & "] where [" & MainField & "] like '" & aID & "'"
      rsNewID.DBOpen strSQL, CNN
      If rsNewID.Recordcount > 0 Then
         x = IIf(IsNull(rsNewID.DBRecordset.Fields(0).Value), 0, rsNewID.DBRecordset.Fields(0).Value) + 1
         newStr = Trim(Str(x))
         z = Len(newStr)
         For Y = 1 To lengthNo - z
            newStr = "0" & newStr
         Next
      Else
         newStr = ""
         For Y = 1 To lengthNo - 1
            newStr = newStr & "0"
         Next
         newStr = newStr & "1"
      End If
      strID = Left(aID, startPos - 1) & newStr & PrefixCounter & Right(aID, Len(aID) - (startPos + lengthNo - 2))
      GetNewID = Replace(strID, "%", "")
   Else
      GetNewID = aID
   End If
End Function

'Private Function GetNewNumber() As String
'   Set rsID = New DBQuick
'   rsID.DBOpen "select "
'End Function

'** fungsi untik mengetahui valid tidak nya sebuah ID **
Public Function IsValidID() As Boolean
   If lNewID = "" Then
      IsValidID = False
   Else
      Dim rsCheck As New DBQuick
      rsCheck.DBOpen "select * from [" & MainTable & "] where [" & MainField & "] = '" & lNewID & "'", CNN
      If rsCheck.Recordcount > 0 Then
         IsValidID = False
      Else
         IsValidID = True
      End If
      Set rsCheck = Nothing
   End If
End Function



'Public Function GetDataLookup() As DBQuick
'   Set rsLookup = New DBQuick
'   Select Case rsFormat.DBRecordset.Fields("type").value
'            Case "Mesin CPP":
'                 rsLookup.DBOpen "Select * from namaTabel", CNN
'            Case "Thickness":
'
'            Case "Film Type":
'
'            Case "Grade SR":
'
'            Case "Posisi GR":
'
'            Case "Posisi Roll":
'   End Select
'   DataLookup = rsLookup
'End Function

