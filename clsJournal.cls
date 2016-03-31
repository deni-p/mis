VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private RcSource As New DBQuick
'Private MyJournal As StrJournalList
Private mIdx, mVarKey As String
Private mVarTR As New clsTransaksi
Private mVarLastJournal As String
'Private Type StrJournalList
'        StrTanggalBukti  As Integer
'        StrNoBukti As Integer
'        StrKodeKaryawan As Integer
'        StrKodeKas As Integer
'        StrKodePartner As Integer
'        StrKodeBarang As Integer
'        StrKodePajak As Integer
'        StrKodeDiscount As Integer
'        StrKodeAktiva As Integer
'        StrKodeHpp As Integer
'        StrDocReff As Integer
'        StrDiscount As Integer
'        StrPpnMasukan As Integer
'        StrPpnKeluaran As Integer
'        StrKasKeluar As Integer
'        StrKasMasuk As Integer
'        StrTotalPembelian As Integer
'        StrPiutangKaryawan As Integer
'        StrPelunasanHutang As Integer
'        StrPerolehanAktiva As Integer
'        StrHPP As Integer
'        strKodeBiaya As Integer
'        strJumlahBiaya As Integer
'        strKodeRetur As Integer
'        strReturPenjualan As Integer
'        strReturPembelian As Integer
'        StrValTanggalBukti  As String
'        StrValNoBukti As String
'        StrValKodeKaryawan As String
'        StrValKodeKas As String
'        StrValKodePartner As String
'        StrValKodeBarang As String
'        StrValKodePajak As String
'        StrValKodeDiscount As String
'        StrValKodeAktiva As String
'        StrValKodeHpp As String
'        StrValDocReff As String
'        StrValPajak As Variant
'        StrValKasKeluar As Variant
'        StrValTotalPembelian As Variant
'        StrValPiutangKaryawan As Variant
'        StrValPelunasanHutang As Variant
'        StrValPerolehanAktiva As Variant
'        StrValDiscount As Variant
'End Type
'
'Public Enum TypeJournalData
'       typJrlBKKPTP = 0
'       typJrlBPBK = 1
'       typJrlBRBP = 2
'       typJrlBKMPTP = 3
'       typJrLBPJK = 4
'       typJrLBRPJ = 5
'End Enum
Private mVarNourut As Integer

'Public Sub CreateJournal(ByVal SourceDataJournal As String, ByVal TipicalJournal As TypeJournalData)
'Dim mVarTypeJrl As String
'If RcSource.DBOpen(SourceDataJournal, Cnn, lckLockReadOnly) = True Then
'   Select Case TipicalJournal
'          Case typJrlBKKPTP: ProsesBacaJournal "BKKPTP"
'          Case typJrlBPBK: ProsesBacaJournal "BPBK"
'          Case typJrlBRBP: ProsesBacaJournal "BRPB"
'          Case typJrlBKMPTP: ProsesBacaJournal "BKMPTP"
'          Case typJrLBPJK: ProsesBacaJournal "BPJK"
'          Case typJrLBRPJ: ProsesBacaJournal "BRPJ"
'   End Select
'   'MessageBox SourceDataJournal
'End If
'End Sub
'
'Private Sub ProsesBacaJournal(ByVal ParamJournal As String)
'Dim RcBcJrl As New DBQuick
''Dim RcSource As New DBQuick
'Dim StrPartic As String
'Dim mTmp As String
'Dim mVarDebet, mVarCredit As Variant
'Dim I As Long
'
'RcBcJrl.DBOpen "SELECT [Daftar Configurasi].NoAccount, [Daftar Configurasi].Posisi, [Daftar Configurasi].[Value Data], [Daftar Configurasi].[Nama Form] FROM  [Tabel Configurasi] INNER JOIN [Daftar Configurasi] ON [Tabel Configurasi].[Kode Konfigurasi] = [Daftar Configurasi].[Kode Konfigurasi] WHERE     ([Tabel Configurasi].[Kode Konfigurasi] = N'" & ParamJournal & "') ORDER BY [Daftar Configurasi].[No Index]", Cnn, lckLockReadOnly
'If RcSource.DBRecordset.Recordcount <> 0 Then
'    mIdx = "JR"
'    TglIndex
'    mTmp = IdxAuto
'    ScanHeaderJournal RcSource.DBRecordset
'    RcSource.DBRecordset.MoveFirst
'    If CiptaKaryaHeaderJournal(mTmp, RcSource.Fields(MyJournal.StrNoBukti), RcSource.Fields(MyJournal.StrDocReff), RcSource.Fields(MyJournal.StrKodeKas), RcSource.Fields(MyJournal.StrKodeKaryawan), RcSource.Fields(MyJournal.StrKodePartner), "IDR", RcSource.Fields(MyJournal.StrTanggalBukti), mVarPeriode, ParamJournal) = True Then
'    'mVarHeader = RcSource.DBRecordset.GetRows(RcSource.DBRecordset.Recordcount, adBookmarkFirst)
'   ' For I = 0 To UBound(mVarHeader, 2)
'        StrPartic = "Transaksi " & mTmp
'        With RcBcJrl.DBRecordset
'             .MoveFirst
'             'mVarTotal = 0
'             Do
'               If .EOF Then Exit Do
'
'                  Select Case UCase(.Fields("Nama Form"))
'                         'Entry Value Data
'                         Case "KAS KELUAR", "PPN MASUKAN", "DISCOUNT", "HUTANG USAHA", "PPN KELUARAN", "PEMBELIAN RETUR", "RETUR PERSEDIAAN", "PENGHASILAN TUNAI", "HPP", "KAS MASUK", "PENJUALAN KREDIT", "PIUTANG USAHA", "PENJUALAN RETUR", "PIUTANG DAGANG":
'                               mVarDebet = HitungTotal(.Fields("Nama Form"))
'                               mVarCredit = mVarDebet
'                               If .Fields("Posisi") = False Then mVarDebet = 0 Else mVarCredit = 0
'                               'mVarHeader(LookDetailFLd(.Fields("Value Data")), I)
'                               StrPartic = StrPartic & "," & .Fields("Nama Form")
'                               CiptaKaryaDetailJournal mTmp, .Fields(0), DetailDocReff(.Fields("Value Data")), CCur(mVarDebet), CCur(mVarCredit)
'                         'Entry Value Data Berdasarkan Barang
'                         Case "TOTAL PEMBELIAN", "TOTAL PERSEDIAAN", "TOTAL PENJUALAN", "TOTAL PENJUALAN BARANG":
'                               RcSource.DBRecordset.MoveFirst
'                               Do
'                                 If RcSource.DBRecordset.EOF Then Exit Do
'                                 mVarDebet = RcSource.DBRecordset.Fields(.Fields("Nama Form").Value)
'                                 mVarCredit = mVarDebet
'                                 If .Fields("Posisi") = False Then mVarDebet = 0 Else mVarCredit = 0
'                                 StrPartic = StrPartic & "," & .Fields("Nama Form")
'                                 CiptaKaryaDetailJournal mTmp, .Fields(0), RcSource.DBRecordset.Fields(.Fields("Value Data").Value), CCur(mVarDebet), CCur(mVarCredit)
'                                 RcSource.DBRecordset.MoveNext
'                               Loop
'                               RcSource.DBRecordset.MoveLast
''                         Case Else
''                               mVarDebet = mVarHeader(LookDetailFLd(.Fields("Nama Form")), I)
''                               mVarCredit = mVarDebet
'                   End Select
'                   .MoveNext
'             Loop
'             .MoveFirst
'             SendDataToServer ("Update [Table Journal] set RefNotes=N'" & Left(StrPartic, 250) & "' where JournalID=N'" & mTmp & "'")
'           End With
'   ' Next I
'    End If
'End If
'RcBcJrl.CloseDB
'RcSource.CloseDB
''MessageBox "Proses Validasi Telah selesai.", "Validasi", msgOkOnly
'End Sub

Public Function CariAkunItem(Params As String) As String
   Dim rsAkun As New DBQuick
   rsAkun.DBOpen "select noAccount from inventory where noItem='" & Params & "'", CNN, lckLockReadOnly
   If rsAkun.DBRecordset.Recordcount > 0 Then
      CariAkunItem = rsAkun.DBRecordset.Fields(0)
   Else
      CariAkunItem = "-"
   End If
   Set rsAkun = Nothing
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
Dim MyData As New clsTransaksi
Dim mVarIndexData As String

If JournalID <> "LINK" Then
   mVarIndexData = MyData.PrepareIndex(tmbTransaksiNOJOURNAL, 13, Format(Year(dDateBegin), "yyyy"), "JR" & Format(Year(dDateBegin), "yyyy") & "-")
Else
   mVarIndexData = "LINK"
End If

CiptaKaryaHeaderJournal = SendDataToServer(" INSERT INTO [Table Journal]" & _
                                           " (JournalID, TransID, PurchaseID, NoAccount, PartnerID,EmpID, Currency, DateTrans,  Periode, TypeTrans,NoUrut)" & _
                                           " VALUES (N'" & mVarIndexData & "', N'" & TransID & "', " & PurchaseID & ", " & BankID & "," & PartnerId & ", " & EmpID & ", N'" & MataUang & "', CONVERT(DATETIME, '" & Format(DateTrans, "dd/mm/yy") & "', 3), " & Periode & ", N'" & TypeTrans & "','" & MyData.PrepareIndex(tmbTransaksiNOJOURNAL, 13, Format(Year(dDateBegin), "yyyy"), "JR" & Format(Year(dDateBegin), "yyyy") & "-") & "')")
mVarLastJournal = mVarIndexData
mVarNourut = 0
Set MyData = Nothing
End Function

Public Sub CreateRefNotes(ByVal RefNotes As String)
SendDataToServer (" Update [Table Journal] Set RefNotes=N'" & Trim(Left(RefNotes, 249)) & "' where JournalID='" & mVarLastJournal & "'")
End Sub

Public Sub CiptaKaryaDetailJournal(ByVal JournalID As String, _
                                    ByVal NoAccount As String, _
                                    ByVal Doc_Reff As String, _
                                    ByVal Debet As Variant, _
                                    ByVal Credit As Variant, _
                                    Optional ByVal Keterangan As String)
If Doc_Reff = "" Then Doc_Reff = "Is Null"
If Keterangan = "" Then Keterangan = "-"
mVarNourut = mVarNourut + 1
SendDataToServer (" INSERT INTO [Detail Journal]" & _
                  " (JournalID, NoAccount, [Doc Reff], Debet, Credit,Keterangan,[No]) " & _
                  " VALUES   (N'" & mVarLastJournal & "', N'" & NoAccount & "', N'" & Doc_Reff & "', " & CCur(Debet) & ", " & CCur(Credit) & ",N'" & Left(Keterangan, 249) & "'," & mVarNourut & ")")
End Sub

Private Function IdxAuto() As String
IdxAuto = mVarTR.PrepareIndex(tmbTransaksiJournal, 5, mVarKey, TglIndex)
End Function

Private Function TglIndex() As String
Dim TglHari As String
Dim TglBulan As String
Dim TglTahun As String
TglIndex = mIdx & "-" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End Function

'Private Sub ScanHeaderJournal(ByVal RecData As Recordset)
'Dim I As Integer
'MyJournal.StrTanggalBukti = 100
'MyJournal.StrNoBukti = 100
'MyJournal.StrKodeKaryawan = 100
'MyJournal.StrKodePartner = 100
'MyJournal.StrKodeKas = 100
'MyJournal.StrKodeBarang = 100
'MyJournal.StrKodePajak = 100
'MyJournal.StrKodeDiscount = 100
'MyJournal.StrKodeAktiva = 100
'MyJournal.StrKodeHpp = 100
'MyJournal.StrDocReff = 100
'MyJournal.StrDiscount = 100
'MyJournal.StrPpnMasukan = 100
'MyJournal.StrPpnKeluaran = 100
'MyJournal.StrKasKeluar = 100
'MyJournal.StrTotalPembelian = 100
'MyJournal.StrPiutangKaryawan = 100
'MyJournal.StrPelunasanHutang = 100
'MyJournal.StrPerolehanAktiva = 100
'MyJournal.StrKasMasuk = 100
'MyJournal.strKodeBiaya = 100
'MyJournal.strJumlahBiaya = 100
'MyJournal.strKodeRetur = 100
'MyJournal.strReturPenjualan = 100
'MyJournal.strReturPembelian = 100
'
'For I = 0 To RecData.Fields.Count - 1
'    Select Case UCase(RecData.Fields(I).Name)
'           Case "TANGGAL BUKTI": MyJournal.StrTanggalBukti = I
'           Case "NO BUKTI": MyJournal.StrNoBukti = I
'           Case "DOC REFF": MyJournal.StrDocReff = I
'           Case "KODE KARYAWAN": MyJournal.StrKodeKaryawan = I
'           Case "KODE CUSTOMER", "KODE SUPPLIER": MyJournal.StrKodePartner = I
'           Case "KODE KAS": MyJournal.StrKodeKas = I
'           Case "KODE BARANG": MyJournal.StrKodeBarang = I
'           Case "KODE PAJAK": MyJournal.StrKodePajak = I
'           Case "KODE DISCOUNT": MyJournal.StrKodeDiscount = I
'           Case "KODE AKTIVA": MyJournal.StrKodeAktiva = I
'           Case "KODE HPP": MyJournal.StrKodeHpp = I
'           Case "PPN MASUKAN": MyJournal.StrPpnMasukan = I
'           Case "PPN KELUARAN": MyJournal.StrPpnKeluaran = I
'           Case "KAS KELUAR": MyJournal.StrKasKeluar = I
'           Case "KAS MASUK": MyJournal.StrKasMasuk = I
'           Case "TOTAL PEMBELIAN": MyJournal.StrTotalPembelian = I
'           Case "PIUTANG KARYAWAN": MyJournal.StrPiutangKaryawan = I
'           Case "PELUNASAN HUTANG": MyJournal.StrPelunasanHutang = I
'           Case "PEROLEHAN AKTIVA": MyJournal.StrPerolehanAktiva = I
'           Case "DISCOUNT": MyJournal.StrDiscount = I
'           Case "HPP": MyJournal.StrHPP = I
'           Case "KODE BIAYA": MyJournal.strKodeBiaya = I
'           Case "KODE RETUR": MyJournal.strKodeRetur = I
'           Case "PEMBELIAN RETUR": MyJournal.strReturPembelian = I
'           Case "JUMLAH BIAYA": MyJournal.strJumlahBiaya = I
'    End Select
'Next I
'End Sub

Private Function LookDetailFLd(ByVal NameStringFld As String) As Integer
Dim i As Integer
For i = 0 To RcSource.DBRecordset.Fields.Count - 1
    If UCase(RcSource.DBRecordset.Fields(i).Name) = UCase(NameStringFld) Then
       LookDetailFLd = i
       Exit For
    End If
Next i
End Function

'Private Function LookDetailJournal(ByVal NameStringFld As String) As Integer
'Select Case UCase(NameStringFld)
'       Case "TANGGAL BUKTI": LookDetailJournal = MyJournal.StrTanggalBukti
'       Case "NO BUKTI": LookDetailJournal = MyJournal.StrNoBukti
'       Case "DOC REFF": LookDetailJournal = MyJournal.StrDocReff
'       Case "KODE KARYAWAN": LookDetailJournal = MyJournal.StrKodeKaryawan
'       Case "KODE CUSTOMER", "KODE SUPPLIER": LookDetailJournal = MyJournal.StrKodePartner
'       Case "KODE KAS": LookDetailJournal = MyJournal.StrKodeKas
'       Case "KODE BARANG": LookDetailJournal = MyJournal.StrKodeBarang
'       Case "KODE PAJAK": LookDetailJournal = MyJournal.StrKodePajak
'       Case "KODE DISCOUNT": LookDetailJournal = MyJournal.StrKodeDiscount
'       Case "KODE AKTIVA": LookDetailJournal = MyJournal.StrKodeAktiva
'       Case "KODE HPP": LookDetailJournal = MyJournal.StrKodeHpp
'       Case "PPN MASUKAN": LookDetailJournal = MyJournal.StrPpnMasukan
'       Case "PPN KELUARAN": LookDetailJournal = MyJournal.StrPpnKeluaran
'       Case "KAS KELUAR": LookDetailJournal = MyJournal.StrKasKeluar
'       Case "KAS MASUK": LookDetailJournal = MyJournal.StrKasMasuk
'       Case "TOTAL PEMBELIAN": LookDetailJournal = MyJournal.StrTotalPembelian
'       Case "PIUTANG KARYAWAN": LookDetailJournal = MyJournal.StrPiutangKaryawan
'       Case "PELUNASAN HUTANG": LookDetailJournal = MyJournal.StrPelunasanHutang
'       Case "PEROLEHAN AKTIVA": LookDetailJournal = MyJournal.StrPerolehanAktiva
'       Case "DISCOUNT": LookDetailJournal = MyJournal.StrDiscount
'       Case "HPP": LookDetailJournal = MyJournal.StrHPP
'       Case "KODE BIAYA": LookDetailJournal = MyJournal.strKodeBiaya
'       Case "KODE RETUR": LookDetailJournal = MyJournal.strKodeRetur
'       Case "PEMBELIAN RETUR": LookDetailJournal = MyJournal.strReturPembelian
'       Case "JUMLAH BIAYA": LookDetailJournal = MyJournal.strJumlahBiaya
'End Select
'End Function

Private Function HitungTotal(ByVal FieldTotal As String) As Variant
Dim mTotal As Variant
Dim mVdata As Variant
Dim RecordsetName As New DBQuick
Dim mVarI As Integer
mTotal = 0
Set RecordsetName.DBRecordset = RcSource.DBRecordset.Clone(adLockReadOnly)
If RecordsetName.Recordcount <> 0 Then
   With RecordsetName.DBRecordset
        mVdata = .Getrows(.Recordcount, adBookmarkFirst, FieldTotal)
        For mVarI = 0 To UBound(mVdata, 2)
            mTotal = mTotal + mVdata(0, mVarI)
        Next mVarI
   End With
End If
HitungTotal = mTotal
RecordsetName.CloseDB
End Function

Private Function DetailDocReff(ByVal FieldTotal As String) As String
Dim mVdata As Variant
Dim RecordsetName As New DBQuick
Dim mVarI As Integer
Set RecordsetName.DBRecordset = RcSource.DBRecordset.Clone(adLockReadOnly)
If RecordsetName.Recordcount <> 0 Then
   With RecordsetName.DBRecordset
        mVdata = .Getrows(.Recordcount, adBookmarkFirst, FieldTotal)
        For mVarI = 0 To UBound(mVdata, 2)
            DetailDocReff = mVdata(0, mVarI)
        Next mVarI
   End With
End If
RecordsetName.CloseDB
End Function

'Private Function NoJrlLink() As String
'Dim Rc As New DBQuick
'Rc.DBOpen "SELECT     MAX(RIGHT(JournalID, 12)) AS MAXNOM FROM         [Table Journal] WHERE     (TypeTrans = N'LINK')", Cnn, lckLockReadOnly
'With Rc
'     If .Recordcount <> 0 Then
'        NoJrlLink = "LINK" & IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + 1
'     Else
'        NoJrlLink = "LINK" & 1
'     End If
'End With
'End Function
