Attribute VB_Name = "Start"
Option Explicit
Public Enum TombolTransaksi
       tmbSupplier = 0
       tmbCustomer = 1
       tmbInventory = 2
       tmbKelompok = 3
       tmbCurrency = 4
       tmbWareHouse = 5
       tmbTransaksiPO = 6
       tmbTransaksiDetailPO = 7
       tmbTransaksiReceive = 8
       tmbTransaksiSC = 9
       tmbTransaksiAR = 10
       tmbVoucher = 11
       tmbShipTransport = 12
       tmbExpedTransport = 13
       tmbDeliveryNotes = 14
       tmbFreight = 15
       tmbAsmOrder = 16
       tmbTransaksiReturBeli = 17
       tmbTransaksiReturjual = 18
       tmbTransaksiMutasiPenjualan = 19
       tmbTransaksiPiutangKaryawan = 20
       tmbTransaksiBayarPiutangKaryawan = 21
       tmbTransaksiBeliAktivaTetap = 22
       tmbTransaksiJualAktivaTetap = 23
       tmbTransaksiJournal = 24
       tmbTransaksiBiayaKeluarJournal = 25
       tmbTransaksiBiayaMasukJournal = 26
       tmbTransaksiBKM = 27
       tmbTransaksiBKK = 28
       tmbTransaksiMutasiGudang = 29
       tmbTransaksiMemorial = 30
       tmbTransaksiinvMemorial = 31
       tmbTransaksiAkumDepre = 32
       tmbTransaksiBKMAT = 33
       tmbTransaksiInvADJ = 34
       tmbTransaksiInvSUB = 35
       tmbTransaksiBKKKARYAWAN = 36
       tmbTransaksiBKMKARYAWAN = 37
       tmbTransaksiHUTANG = 36
       tmbTransaksiPIUTANG = 37
       tmbTransaksiBKKAT = 38
       tmbTransaksiNOJOURNAL = 39
       tmbTransaksiChange = 40
End Enum

Public Enum TypePO
       PORDER = 1
       PNORM = 2
       PMRP = 3
       PCASH = 4
       PBLANKET = 5
End Enum

Public Enum TypeSO
       SORDER = 1
       SCONTRACT = 2
       SQUOTE = 3
End Enum

Public Enum MessageEnumBox
       msgRetryYesNo = 0
       msgYesNo = 1
       msgNo = 2
       msgOkOnly = 3
End Enum

Public Enum MessageManner
       msgCrtical = 0
       msgInfo = 1
       msgQuestion = 2
       msgExclamation = 3
End Enum

Public Enum ButtonTransDB
       BtnNone = 0
       BtnEdit = 1
       BtnAddnew = 2
       BtnDelete = 3
End Enum
Public Enum ModeCheckDate
   TimeData
   DateData
End Enum
Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Type SetingJournalData
       JournalFieldString As String
       JournalValueString As String
       Posisijournal As Byte
End Type


Public IsLogOff As Boolean
Public aksess As New clsUserAccess
Public m_Report As New CRAXDRT.Report
Public frmSPPH As FrmPurchaseOffer
Public frmSPH As FrmPurchaseOffer
Public CallerID As Boolean
Public mVarSetingJournalData As SetingJournalData
Public KuBox As Integer
Public dDateBegin, dDateEnd, StrCnn, JournalFieldString, mVarLoginActive, JournalValueString As String
Public CNN As New Connection
Public mVarPeriode, mVarTempPeriode, mVarIDUser As Integer
Public IsLoginSucces As Boolean
Public IsFrmSup, IsFrmCus, IsFrmPo, IsfrmSc As Boolean
Public TahunFiskalYear, mVarPassword, mVarServerName, mVarUserID As String
Public mUserName As String
Public Tsample As Boolean   'digunakan di Permintaan sample
Public TSalesforcast As Boolean 'digunakan di sales forcast



Private mVarADODC As String
Public Isregis, IsEnabledLogin As Boolean

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8

Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Private Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const WS_CHILDWINDOW = &H40000000
Private Const GWL_STYLE = (-16)
Private Const HWND_TOPMOST = -1
Private Const WS_CHILD = &H40000000
Private Const WS_POPUP = &H80000000
Private Const GWL_HWNDPARENT = -8
Private Const WS_CAPTION = &HC00000
Private Const WS_THICKFRAME = &H40000
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_FORCE_REDRAW = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER
Private Const HWND_NOTOPMOST = -2
Private Const WM_ACTIVATE As Long = &H6
Public ReportPath As String
Public ReportPos As String
Public FirstNode As String

'variabel enkripdata
Private sBoxInit(0 To 15)   As Variant
Private BoxTurnOver         As Variant
Private sBoxPos(32)         As Byte
Private sBox(32)            As Byte
Private sBoxOut(32)         As Byte
Private sBoxInvInit(32, 15) As Byte

Private aDecTab(255)        As Integer
Private aEncTab(63)         As Byte

Public Key As String


Declare Function ShellExecute Lib "shell32.dll" Alias _
   "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
   As String, ByVal lpFile As String, ByVal lpParameters _
   As String, ByVal lpDirectory As String, ByVal nShowCmd _
   As Long) As Long

Declare Function apiFindWindow Lib "user32" Alias "FindWindowA" _
   (ByVal lpclassname As Any, ByVal lpCaption As Any) As Long

Global vTransaction As ADODB.Recordset


Global Const SW_SHOWNORMAL = 1
Public IDLELIMIT As String   'On second
Public StartFromIdle As Boolean
Public CurrentDept, NamaDept, GroupID As String

Private Const LOCALE_SDATE = &H1F
Private Const LOCALE_STIMEFORMAT = &H1003

Private Const WM_SETTINGCHANGE = &H1A

Private Const HWND_BROADCAST = &HFFFF&

Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Public fmtBooleanData As StdDataFormat
Public fmtBoolOpenClose As StdDataFormat
Global Const RecForm = "#,##0"
Global Const QtyForm = "#,##0;(#,##0)"
Global Const QtyFormFloat = "#,##0.000;(#,##0.000)"
Global Const AngkaForm = "#.##0,00;(#.##0,00)"
Global Const PriceForm = "#,##0;(#,##0)" '"#,##
Global Const ShortDateForm = "dd MMM yyyy"
Global Const ShortDateFormGaris = "dd-MMM-yyyy"
Global Const ProcentForm = "0%;(0%)"
Global Const LongDateForm = "dd/mm/yy hh:mm:ss"
Public fmtBoolActivate As StdDataFormat
Global AINode As Node
Global FirstTgl, LastTgl As Date
Global StartTgl, EndTgl As Date


'******* MENU TAMBAHAN ****************

Public mnBInventoryAdj As Boolean
Public mnBInventoryBrowser As Boolean
Public mnStockOpname As Boolean

Public mnAPPPermintaanSample As Boolean
Public mnAPPMemo As Boolean
Public mnAPPCustomerFeedback As Boolean

Public mnAPPEvaSupp As Boolean
Public mnAPPSuratJalan As Boolean
Public mnAPPSuratRetur As Boolean
Public mnBApprovalPO As Boolean
Public mnAPPPermintaanHarga As Boolean
Public mnAPPRKPPermintaanBeli As Boolean
Public mnAPPSPB As Boolean
Public mnAPPLPB As Boolean
Public mnAPPRPBRL As Boolean
Public mnAPPTTRL As Boolean   'app tanda terima RL
Public mnAPPLembSupp As Boolean   'app lembar supplier
Public mnAPPKRL As Boolean    ' app pengiriman rl

Public mnAPPALKALI As Boolean
Public mnAPPACID As Boolean
Public mnAPPBLEACHING As Boolean
Public mnAPPEKSREAKTOR As Boolean
Public mnAPPEKSAUTO As Boolean
Public mnAPPFILTERPRESS As Boolean
Public mnAPPGELL As Boolean
Public mnAPPBUNGKUS As Boolean
Public mnAPPCONCRETE As Boolean
Public mnAPPHYDRAULIC As Boolean
Public mnAPPCUTTER As Boolean
Public mnAPPJEMUR As Boolean
Public mnAPPCRUSHER As Boolean
Public mnAPPMIXING As Boolean
Public mnAPPBLENDING As Boolean
Public mnAPPSTFG As Boolean
Public mnAPPPAKAI As Boolean

'********    END   ********************


Public Function IsPersediaanReady() As Boolean

Dim rsCheck As New DBQuick
Dim I As Integer
Dim Avdata As Variant
Dim strSQL As String

strSQL = "SELECT GLAccount.NoAccount AS Kode_GL, [Inventory Group].NoAccount AS Kode_Group " & _
        " FROM GLAccount RIGHT OUTER JOIN [Inventory Group] ON GLAccount.NoAccount = [Inventory Group].NoAccount " & _
        " Where (GLAccount.noAccount Is Null)"
rsCheck.DBOpen strSQL, CNN, lckLockReadOnly

With rsCheck
     If .Recordcount <> 0 Then
        IsPersediaanReady = False
     Else
        IsPersediaanReady = True
     End If
End With
rsCheck.CloseDB
End Function


Public Function SetDateTime() As Boolean
   Dim dwLCID As Long
   dwLCID = GetSystemDefaultLCID()
   
   If SetLocaleInfo(dwLCID, LOCALE_SDATE, "dd/MM/yy") = False Then
   SetDateTime = False
   Exit Function
   End If
   
   If SetLocaleInfo(dwLCID, LOCALE_STIMEFORMAT, "HH:mm:ss") = False Then
   SetDateTime = False
   Exit Function
   End If
   
   PostMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
   
   SetDateTime = True
End Function



Public Sub Main()
IsLogOff = False
CheckUpdate

SetDateTime
StartFromIdle = False
Key = "BulirMMTPadi74"  'digunakan untuk key enkripsi

'**** init variabel sementara *******
ReportPath = "E:\Manufacture\Report"
ReportPos = ReportPath
'************************************'



On Error Resume Next
'mVarPeriode = 1
'TahunFiskalYear = Format(Date, "yyyy")
If MakeAdoDC = False Then
     If Isregis = False Then
        MessageBox "Anda Belum Mengisi Data Perusahaan Anda.", "Perusahaan", "msgOkOnly", msgExclamation
        MainMenu.Show
        frmAbout.SSTab1.Tab = 1
        frmAbout.SetFocus
     Else
         dDateBegin = Format(Date, "dd/mm/yyyy")
         dDateEnd = Format(Date, "dd/mm/yyyy")
'            MainMenu.Show
         frmLogin.Show vbModal
         MainMenu.Show
         'MainMenu.SemeruTree1.Visible = True
         If IsLoginSucces = True Then
            If OpenCnn(StrCnn) = True Then
               If PeriodeBerjalan = False Then
                  'FrmSetingPeriode.SetFocus
               Else
                  MainMenu.Show
               End If
            Else
            End If
         Else
            End
         End If
     End If
Else
   MessageBox "Silahkan anda melakukan registrasi kepada Bulirpadi Lintas Nusantara ,PT", "Registrasi", msgOkOnly, msgExclamation
End If

Set fmtBooleanData = New StdDataFormat
fmtBooleanData.Type = fmtBoolean
fmtBooleanData.TrueValue = "YES"
fmtBooleanData.FalseValue = "NO"
fmtBooleanData.NullValue = ""

Set fmtBoolOpenClose = New StdDataFormat
fmtBoolOpenClose.Type = fmtBoolean
fmtBoolOpenClose.TrueValue = "CLOSED"
fmtBoolOpenClose.FalseValue = "OPEN"
fmtBoolOpenClose.NullValue = ""


Set fmtBoolActivate = New StdDataFormat
fmtBoolActivate.Type = fmtBoolean
fmtBoolActivate.TrueValue = "AKTIF"
fmtBoolActivate.FalseValue = "NON-AKTIF"
fmtBoolActivate.NullValue = ""


If Err.Number <> 0 Then MsgBox Err.Number & vbCrLf & Err.Description
Err.Clear
End Sub


Public Function SQLLookupParameter(aRecordSet As ADODB.Recordset, aLookupField As String, _
                                    valueField As String, Optional AdditionalParams As String = "") As String
Dim strParams As String
   strParams = IIf(Trim(AdditionalParams) = "", "", " where (" & AdditionalParams & ")")
   If aRecordSet.Recordcount > 0 Then
      With aRecordSet
         .MoveFirst
         While Not .EOF
            strParams = strParams + IIf(Trim(strParams) = "", " where ", " and ") + _
                        "[" & aLookupField & "] <> '" & .Fields(valueField) & "'"
            .MoveNext
         Wend
      End With
   End If
   SQLLookupParameter = strParams
End Function



Public Function OpenCnn(Optional ByVal ConnectionStringDB As String) As Variant
On Error GoTo Hell
If Not CNN Is Nothing Then
   If CNN.State = 0 Then
      CNN.CursorLocation = adUseClient
      CNN.IsolationLevel = adXactChaos
      CNN.Mode = adModeShareDenyNone
      CNN.ConnectionString = ConnectionStringDB
      CNN.Open
      OpenCnn = True
   ElseIf CNN.State = 1 Then
      OpenCnn = True
   End If
Else
'   If Cnn.State = 0 Then
'      Cnn.CursorLocation = adUseClient
'      Cnn.IsolationLevel = adXactChaos
'      Cnn.Mode = adModeShareDenyNone
'      CNN = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Bilinus Man;Data Source=BULIRCOMP"
'      Cnn.Open
'      OpenCnn = True
   End If
'End If
Exit Function
Hell:
   OpenCnn = Err.Number
   'MsgBox Err.Description
   MessageBox "Login Belum Sukses. Harap Diulangi Sekali Lagi." & vbCrLf & vbCrLf & "Open Connection With Reason:" & vbCrLf & Err.Number & " - " & Err.Description, "Login Failed", msgOkOnly
   Err.Clear
End Function

Public Sub CloseDB(ByRef RcActive As Recordset)
If Not RcActive Is Nothing Then
   If RcActive.State = 1 Then
      RcActive.Close
   End If
End If
Set RcActive = Nothing
End Sub

Public Sub Block(ByRef MyObj As TextBox)
On Error Resume Next
       MyObj.SelStart = 0
       MyObj.SelLength = Len(MyObj)
Err.Clear
End Sub

Public Sub KeyEnter(ByVal KeyCode As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Public Function KirimNull(ByVal LenIdx As Integer) As String
Dim I As Integer
KirimNull = ""
For I = 1 To LenIdx
    KirimNull = Trim(KirimNull & "0")
Next I
End Function


Public Function ScanGrid(ByRef DataGridActive As DataGrid, Optional ByVal NoColumn As String) As Boolean
On Error GoTo Hell
Dim I As Integer
With DataGridActive
     For I = 0 To .Columns.Count - 1
         If .Columns(I).Visible = True Then
            If NoColumn <> "" Then
                If .Columns(Val(NoColumn)).Value = "0" Then
                   MessageBox "Data Tidak Boleh Kosong........!Harus Bernilai.....!", "Peringatan", msgOkOnly
                   ScanGrid = True
                   Exit Function
                End If
                If .Columns(I).Value = "" Then
                   MessageBox "Data Tidak Boleh Kosong........!", "Peringatan", msgOkOnly
                   ScanGrid = True
                   Exit Function
                End If
            Else
                If .Columns(I).Value = "" Then
                   MessageBox "Data Tidak Boleh Kosong........!", "Peringatan", msgOkOnly
                   ScanGrid = True
                   Exit Function
                End If
                If .Columns(I).Value = "0" Then
                   MessageBox "Data Tidak Boleh Kosong........!Harus Bernilai.....!", "Peringatan", msgOkOnly
                   ScanGrid = True
                   Exit Function
                 End If
            End If
         End If
     Next I
End With
Exit Function
Hell:
     MessageBox "Data Tidak Boleh Kosong........!", "Peringatan", msgOkOnly
     Err.Clear
End Function

Public Function MessageBox(ByVal StringMessage As String, Optional TitelString As String, Optional MessageBoxEnum As MessageEnumBox = msgOkOnly, Optional MessageLogo As MessageManner) As Integer
On Error GoTo Hell
KuBox = 0
Dim fmess As New frmMessage
With fmess
     If TitelString = "" Then TitelString = "Warning"
     .lblMessage = TitelString
     .txtMessage = StringMessage
     Select Case MessageBoxEnum
            Case msgRetryYesNo:
            Case msgYesNo:
                 .cmdOk(1).Visible = True
                 .cmdOk(1).Enabled = True
                 .cmdOk(1).Caption = "&Ya"
                 '.SetFocus
                 .cmdOk(2).Visible = True
                 .cmdOk(2).Enabled = True
                 .cmdOk(2).Caption = "&Tidak"
'                 .ImAsk.Visible = True
'                 .ImAsk.ZOrder 0
                Select Case MessageLogo
                   Case msgCrtical
                       .ImCritical.Visible = True
                       .ImCritical.ZOrder 0
                   Case msgExclamation
                       .ImExclamation.Visible = True
                       .ImExclamation.ZOrder 0
                   Case msgInfo
                       .ImInfo.Visible = True
                       .ImInfo.ZOrder 0
                   Case msgQuestion
                       .ImQuestion.Visible = True
                       .ImQuestion.ZOrder 0
                End Select
            Case msgNo:
            Case msgOkOnly:
                 .cmdOk(2).Visible = True
                 .cmdOk(2).Enabled = True
                 .cmdOk(2).Caption = "&Lanjut"
                Select Case MessageLogo
                   Case msgCrtical
                       .ImCritical.Visible = True
                       .ImCritical.ZOrder 0
                   Case msgExclamation
                       .ImExclamation.Visible = True
                       .ImExclamation.ZOrder 0
                   Case msgInfo
                       .ImInfo.Visible = True
                       .ImInfo.ZOrder 0
                   Case msgQuestion
                       .ImQuestion.Visible = True
                       .ImQuestion.ZOrder 0
                End Select
    End Select
End With
fmess.Show vbModal
Set fmess = Nothing
MessageBox = KuBox
Exit Function
Hell:
    MsgBox "MessageBox Error" & vbCrLf & Err.Description, , "Warning MessageBox"
    Err.Clear
End Function
Public Sub SendVoucher(ByVal TransID As String, _
                       ByVal PartnerId As String, _
                       ByVal RefNotes As String, _
                       ByVal DateTrans As String, _
                       ByVal Debet As Variant, _
                       ByVal Credit As Variant, _
                       ByVal PurchaseID As Variant, _
                       ByVal TypeTrans As String)
Dim mTrans As New clsTransaksi
With mTrans
     .PrepareVoucher TransID, PartnerId, RefNotes, DateTrans, Debet, Credit, PurchaseID, TypeTrans
End With
Set mTrans = Nothing
End Sub

Public Sub SendAPItem(ByVal NoItem As String, _
                      ByVal QTY_IN As Double, _
                      ByVal PriceIn As Currency, _
                      ByVal RefTrans As String, _
                      ByVal DateTrans As String, _
                      ByVal TypeTrans As String, _
                      ByVal Discount As Single, _
                      ByVal PPN As Single, Optional DeleteFirst As Boolean = False, _
                      Optional wHouse As String)
Dim mTrans As New clsTransaksi
With mTrans
     If DeleteFirst = True Then
        SendDataToServer ("Delete From [Inventory Tabel] where NoItem=N'" & NoItem & "' and RefTrans =N'" & RefTrans & "' and lokasiGdg='" & wHouse & "'")
     End If
     .PrepareAPItem NoItem, QTY_IN, PriceIn, RefTrans, DateTrans, TypeTrans, Discount, PPN, wHouse
End With
Set mTrans = Nothing
End Sub


Public Sub SendARItem(ByVal NoItem As String, _
                      ByVal QTY_OUT As Variant, _
                      ByVal PriceOut As Currency, _
                      ByVal RefTrans As String, _
                      ByVal DateTrans As String, _
                      ByVal HppBarang As Variant, _
                      ByVal TypeTrans As String, Optional DeleteFirst As Boolean = False)
Dim mTrans As New clsTransaksi
With mTrans
     If DeleteFirst = True Then
        SendDataToServer ("Delete From [Inventory Tabel] where NoItem=N'" & NoItem & "' and RefTrans =N'" & RefTrans & "'")
     End If
     .PrepareARItem NoItem, QTY_OUT, PriceOut, RefTrans, DateTrans, HppBarang, TypeTrans
End With
Set mTrans = Nothing
End Sub

Public Sub CallRPTReport(ByVal ReportFileName As String, Optional ByVal SQLQueryString As String, Optional ByVal UkuranKertas As Boolean = False, Optional ByVal KonversiMataUang As Currency, Optional ByVal qryOpenSubReport As String, Optional ByVal SubReportName As String)
On Error GoTo Hell
Screen.MousePointer = vbHourglass
Dim RcRpt As New Recordset
Dim RcTes As New Recordset
Dim RcSub As New Recordset

With RcRpt
     .CursorLocation = adUseClient
     .Open "SELECT FileNameReport, [Alias Report], ViewObject FROM [Report Modules] WHERE (FileNameReport = N'" & ReportFileName & "')", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
End With
If RcRpt.Recordcount <> 0 Then
   RcTes.CursorLocation = adUseClient
   RcTes.Open " Select * from [" & UCase(RcRpt.Fields("ViewObject").Value) & "]", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   If RcTes.Recordcount <> 0 Then
      Dim Mprint As New frmReportView
      With Mprint
           If SQLQueryString = "" Then
              .CallReport "Select * from [" & UCase(RcRpt.Fields("ViewObject").Value) & "]", ReportFileName, ReportPath, RcRpt.Fields("Alias Report")
           Else
              .CallReport SQLQueryString, ReportFileName, ReportPath, RcRpt.Fields("Alias Report").Value
           End If
'           If KonversiMataUang <> 0 Then .m_Report.ReportAuthor = "Terbilang : " & Konversi(KonversiMataUang)
           If SubReportName <> "" Then .SubReport1 qryOpenSubReport, SubReportName
           .SetFocus
      End With
   Else
      MessageBox "Laporan Belum Siap.......!" & vbCrLf & "Atau Laporan Belum Ada Datanya!", "Laporan", msgOkOnly, msgExclamation
   End If
   CloseDB RcRpt
   CloseDB RcTes
Else
   MessageBox "Laporan Belum Siap.......!" & vbCrLf & "Atau Laporan Belum Ada Datanya!", "Laporan", msgOkOnly, msgExclamation
End If
Screen.MousePointer = vbDefault
Exit Sub
Hell:
    MessageBox Err.Description, "Peringatan", msgOkOnly, msgExclamation
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub

Public Sub CloseAllForm()
Dim I As Integer
On Error GoTo Hell
For I = 0 To Forms.Count
    If Screen.ActiveForm.Tag <> "MAIN" Then
       Unload Screen.ActiveForm

    End If
Next I
Hell:
    Err.Clear
End Sub

Public Function SeekForm(ByVal FormCaption As String) As Boolean
Dim I As Integer
Dim frm As Form
On Error GoTo Hell

For Each frm In Forms
    If UCase(frm.Caption) = UCase(FormCaption) Then
       SeekForm = True
       frm.ZOrder (0)
    End If
Next
Set frm = Nothing
Hell:
    Err.Clear
End Function

Public Sub ValidNum(ByRef Key As Integer)
Dim StrValid As String
StrValid = "0123456789-.,"
If Key > 26 Then
   If InStr(StrValid, Chr(Key)) = 0 Then
      Key = 0
   End If
End If
End Sub

Public Function FindOwnRecordset(ByRef RecordsetName As Recordset, ByVal GetFindString As String) As Boolean
Dim mNewcari As Recordset
Set mNewcari = RecordsetName.Clone(adLockReadOnly)
With mNewcari
     If .Recordcount <> 0 Then
        .Filter = GetFindString
        If .Recordcount >= 2 Then
           FindOwnRecordset = True
        Else
           FindOwnRecordset = False
        End If
     Else
        FindOwnRecordset = False
     End If
End With
CloseDB mNewcari
End Function

Public Function ValidString(ByVal QueryFields As String) As String
ValidString = Replace(QueryFields, ",", "")
ValidString = Replace(ValidString, "'", "")
End Function

Public Function SendDataToServer(ByVal SendQueryToServer As String) As Boolean
Dim iCn  As New Connection
Dim Icom As New Command
On Error GoTo Err
'    Debug.Print SendQueryToServer
   With iCn
    .CursorLocation = adUseClient
    .Mode = adModeShareExclusive
    .IsolationLevel = adXactChaos
    .ConnectionString = CNN
    .Open
    If .State = 1 Then
       If SendQueryToServer <> "" Then
          Set Icom.ActiveConnection = iCn
          Icom.CommandType = adCmdText
          Icom.CommandText = SendQueryToServer
'          MsgBox SendQueryToServer
          Icom.Execute
          Set Icom = Nothing
          SendDataToServer = True
       End If
    End If
   End With
   iCn.Close
   Set iCn = Nothing
Exit Function
Err:
     If Not iCn Is Nothing Then
        If iCn.State = 1 Then
           iCn.Close
        End If
     End If
     Set iCn = Nothing
     MessageBox "Operation Failed........!" & vbCrLf & "Chek This Query ->" & vbCrLf & SendQueryToServer & vbCrLf & "With Reason Number :" & Err.Number & vbCrLf & "With Description : " & Err.Description, "Warning SendataToServer", msgOkOnly, msgExclamation
     Err.Clear
End Function

Public Sub ScanKey(ByVal VarKeycode As Integer, ByVal VarShift As Integer, ByRef SetControlData As SemeruOleDC)
On Error GoTo KeyErr
If VarShift = 0 Then
    Select Case VarKeycode
       Case vbKeyF2:
            SetControlData.CallButtonActive tmbEdit
       Case vbKeyF3:
            SetControlData.CallButtonActive tmbAddNew
       Case vbKeyF4:
            SetControlData.CallButtonActive tmbDelete
       Case vbKeyF5:
            SetControlData.CallButtonActive tmbSave
       Case vbKeyF6:
            SetControlData.CallButtonActive tmbCancel
       Case vbKeyF7:
            SetControlData.CallButtonActive tmbTools
       Case vbKeyF8:
            SetControlData.CallButtonActive tmbPrint
       Case vbKeyF9:
            SetControlData.CallButtonActive tmbDetail
       Case vbKeyF10, vbKeyEscape:
            SetControlData.CallButtonActive tmbQuit
       Case vbKeyF11:
            SetControlData.CallButtonActive tmbNextRecord
       Case vbKeyF12:
            SetControlData.CallButtonActive tmbPreviousRecord
       Case Else
          Exit Sub
    End Select
ElseIf VarShift = 2 Then
    Select Case VarKeycode
       Case vbKeyF11:
            SetControlData.CallButtonActive tmbTopRecord
       Case vbKeyF12:
            SetControlData.CallButtonActive tmbBottomRecord
       Case Else
          Exit Sub
    End Select
End If
Exit Sub
KeyErr:
   MsgBox Err.Description, vbCritical, Screen.ActiveForm.Name & "-KeyDown"
   Err.Clear
End Sub

Public Sub ScanKeyGrid(ByVal VarKeycode As Integer, ByVal VarShift As Integer, ByRef SetControlData As SemeruOleDC)
On Error GoTo KeyErr
If VarShift = 0 Then
    Select Case VarKeycode
       Case vbKeyF2:
            SetControlData.CallButtonActive tmbEdit
       Case vbKeyF3:
            SetControlData.CallButtonActive tmbAddNew
       Case vbKeyF4:
            SetControlData.CallButtonActive tmbDelete
       Case vbKeyF5:
            SetControlData.CallButtonActive tmbSave
       Case vbKeyF6:
            SetControlData.CallButtonActive tmbCancel
       Case vbKeyF7:
            SetControlData.CallButtonActive tmbTools
       Case vbKeyF8:
            SetControlData.CallButtonActive tmbPrint
       Case vbKeyF9:
            SetControlData.CallButtonActive tmbDetail
       Case vbKeyF10, vbKeyEscape:
            SetControlData.CallButtonActive tmbQuit
       Case vbKeyF11:
            SetControlData.CallButtonActive tmbNextRecord
       Case vbKeyF12:
            SetControlData.CallButtonActive tmbPreviousRecord
       Case Else
          Exit Sub
    End Select
ElseIf VarShift = 2 Then
    Select Case VarKeycode
       Case vbKeyF11:
            SetControlData.CallButtonActive tmbTopRecord
       Case vbKeyF12:
            SetControlData.CallButtonActive tmbBottomRecord
       Case Else
          Exit Sub
    End Select
End If
Exit Sub
KeyErr:
   MsgBox Err.Description, vbCritical, Screen.ActiveForm.Name & "-KeyDown"
   Err.Clear
End Sub

'Public Function OpenUserAccess() As Recordset
'Dim Icom As New Command
'With Icom
'     .ActiveConnection = Cnn
'     .CommandType = adCmdStoredProc
'     .CommandText = "UserView"
'     Set OpenUserAccess = .Execute
'End With
'End Function

'Public Function ItemSinkron() As Boolean
'On Error GoTo Hell
'Dim cmdCmm As New Command
'Dim Cnnx As New Connection
'Cnnx.CursorLocation = adUseClient
'Cnnx.Mode = adModeShareExclusive
'Cnnx.IsolationLevel = adXactChaos
'Cnnx.ConnectionString = Cnn.ConnectionString
'Set cmdCmm.ActiveConnection = Cnnx
'cmdCmm.CommandType = adCmdText
'cmdCmm.CommandText = " DELETE FROM [Inventory Tabel]"
'cmdCmm.Execute
'
'cmdCmm.CommandText = " INSERT INTO [Inventory Batch]" & _
'                     " (NoItem, QTY_IN, PriceIn, QTY_OUT, PriceOut)" & _
'                     " SELECT NoItem, SUM(QTY_IN) AS QTY_IN, AVG(PriceIn) AS PriceIn, SUM(QTY_OUT) AS QTY_OUT, MAX(PriceOut) AS PriceOut FROM [Inventory Tabel] GROUP BY NoItem ORDER BY NoItem"
'cmdCmm.Execute
'Set cmdCmm = Nothing
'Cnnx.Close
'Set Cnnx = Nothing
'ItemSinkron = True
'Exit Function
'Hell:
'    Err.Clear
'End Function

Public Function BoolToInt(ByVal SetFieldData As Boolean) As Integer
If SetFieldData = True Then BoolToInt = 1 Else BoolToInt = 0
End Function

Public Function MakeAdoDC() As Boolean
MakeAdoDC = False
Isregis = True
Exit Function
Dim MyEx As New clsTemp
Dim mStTemp As String
Dim mCmp, mVarIsBroken As String
Dim mCmp1, mCmp2, mCmpTglAkhir As String
mStTemp = GetSetting(App.EXEName, "Lisence Profile", "Company Name")
mStTemp = mStTemp & GetSetting(App.EXEName, "Lisence Profile", "Address")
mStTemp = mStTemp & GetSetting(App.EXEName, "Lisence Profile", "City")
mStTemp = mStTemp & GetSetting(App.EXEName, "Lisence Profile", "Phone")
mStTemp = mStTemp & GetSetting(App.EXEName, "Lisence Profile", "NPWP")

If UCase(MyEx.SlowRemuse(GetSetting(App.EXEName, "Lisence Profile", "SID"), UCase(mStTemp))) = "BROKEN RULES" Then
   MsgBox "File Corruption Detected.", vbInformation + vbOKOnly, "File Corruption Detected"
   End
End If

If mStTemp <> "" And MyEx.FileExist(App.Path & "\" & App.EXEName & ".Arj") = True Then
   mCmpTglAkhir = Trim(Right(MyEx.SlowRemuse(MyEx.BukafilePath(App.Path & "\" & App.EXEName & ".Arj"), UCase(mStTemp)), 10))
   If IsDate(mCmpTglAkhir) = False Then
      MsgBox "Time To Use This Application Have Been Done." & vbCrLf & "Thanks To Evaluated My Software." & vbCrLf & "Please Contact This Software Vendor For More Information" & vbCrLf & vbCrLf & "Bulirpadi Lintas Nusantara PT" & vbCrLf & "Phone : 0341-470168" & vbCrLf & "Email : Bulirpadi@Bilinus.Com", vbInformation + vbOKOnly, "Evaluation Time"
      End
   End If
Else
   If mStTemp <> "" Then
      MsgBox "File Corruption Detected.", vbInformation + vbOKOnly, "File Corruption Detected"
      SaveSetting App.EXEName, "Lisence Profile", "SID", MyEx.FlashRemuse("BROKEN RULES", UCase(mStTemp))
      If MyEx.FileExist(App.Path & "\" & App.EXEName & ".Arj") = True Then Kill App.Path & "\" & App.EXEName & ".Arj"
      End
      
   End If
End If
If mStTemp <> "" Then
    Isregis = True
    mCmp = MyEx.SlowRemuse(GetSetting(App.EXEName, "Lisence Profile", "SID"), mStTemp)
    'Tgl Install
    mCmp1 = MyEx.SlowRemuse(GetSetting(App.EXEName, "Lisence Profile", "CLSID CMP"), mStTemp)
    'Rentang Run App
    mCmp2 = MyEx.SlowRemuse(GetSetting(App.EXEName, "Lisence Profile", "CLSID DMN"), mStTemp)
    'MsgBox CDate(Format(mCmp2, "dd/mm/yyyy"))
    If mCmp1 = "" And mCmp2 = "" Then
       Isregis = False
    Else
        'MsgBox Date & " - " & CDate(Format(mCmpTglAkhir, "dd/mm/yyyy"))
'        MsgBox CDate(Format(Date, "dd/mm/yyyy")) & " - " & CDate(Format(mCmpTglAkhir, "dd/mm/yyyy"))
        If Date < CDate(Format(mCmpTglAkhir, "dd/mm/yy")) Then
           'Back Date
           MsgBox "Back Date Period Has Been Detected." & vbCrLf & vbCrLf & "Can't Back Date Period In Trial Application." & vbCrLf & "Please Contact This Software Vendor For More Information" & vbCrLf & vbCrLf & "Bulirpadi Lintas Nusantara PT" & vbCrLf & "Phone : 0341-470168" & vbCrLf & "Email : Bulirpadi@Bilinus.Com", vbInformation + vbOKOnly, "Back Date Detected"
           End
        ElseIf Date > CDate(Format(mCmpTglAkhir, "dd/mm/yy")) Then
           'Counter
           If (CDate(Format(mCmp2, "dd/mm/yyyy")) - CDate(Format(Date, "dd/mm/yyyy"))) > 0 Then
              MessageBox "Waktu Anda Untuk Menggunakan Aplikasi Ini Tinggal " & CDate(Format(mCmp2, "dd/mm/yyyy")) - Date & " Hari.", "Registrasi", msgOkOnly, msgInfo
              MyEx.SimpanfilePath App.Path & "\" & App.EXEName & ".Arj", UCase(mStTemp), Format(Date, "dd/mm/yyyy")
              SaveSetting App.EXEName, "Lisence Profile", "CLSID CMP", MyEx.FlashRemuse(Format(Date, "dd/mm/yyyy"), UCase(mStTemp))
              MakeAdoDC = False
           Else
              MsgBox "Time To Use This Application Have Been Done." & vbCrLf & "Thanks To Evaluated My Software." & vbCrLf & "Please Contact This Software Vendor For More Information" & vbCrLf & vbCrLf & "Bulirpadi Lintas Nusantara PT" & vbCrLf & "Phone : 0341-470168" & vbCrLf & "Email : Bulirpadi@Bilinus.Com", vbInformation + vbOKOnly, "Evaluation Time"
              MyEx.SimpanfilePath App.Path & "\" & App.EXEName & ".Arj", UCase(mStTemp), "END DISKAS"
              SaveSetting App.EXEName, "Lisence Profile", "CLSID CMP", MyEx.FlashRemuse("END DISKAS", UCase(mStTemp))
              MakeAdoDC = False
              End
           End If
        ElseIf CDate(Format(Date, "dd/mm/yyyy")) = CDate(Format(mCmpTglAkhir, "dd/mm/yyyy")) Then
           If (CDate(Format(mCmp2, "dd/mm/yyyy")) - CDate(Format(Date, "dd/mm/yyyy"))) > 0 Then
              If CDate(Format(mCmp2, "dd/mm/yyyy")) - CDate(Format(Date, "dd/mm/yyyy")) <> 0 Then
                 MessageBox "Waktu Anda Untuk Menggunakan Aplikasi Ini Tinggal " & CDate(Format(mCmp2, "dd/mm/yyyy")) - CDate(Format(Date, "dd/mm/yyyy")) & " Hari.", "Registrasi", msgOkOnly, msgInfo
                 MyEx.SimpanfilePath App.Path & "\" & App.EXEName & ".Arj", UCase(mStTemp), Format(Date, "dd/mm/yyyy")
                 SaveSetting App.EXEName, "Lisence Profile", "CLSID CMP", MyEx.FlashRemuse(Format(Date, "dd/mm/yyyy"), UCase(mStTemp))
                 MakeAdoDC = False
              End If
           Else
              MessageBox "Waktu Anda Untuk Menggunakan Aplikasi Ini Tinggal " & CDate(Format(mCmp2, "dd/mm/yyyy")) - CDate(Format(Date, "dd/mm/yyyy")) & " Hari.", "Registrasi", msgOkOnly, msgInfo
              MyEx.SimpanfilePath App.Path & "\" & App.EXEName & ".Arj", UCase(mStTemp), "END DISKAS"
              SaveSetting App.EXEName, "Lisence Profile", "CLSID CMP", MyEx.FlashRemuse("END DISKAS", UCase(mStTemp))
              MakeAdoDC = False
           End If
        End If
    End If
Else
   Isregis = False
End If
Set MyEx = Nothing
End Function

Private Sub ColForm(ByVal BoxContainer As PictureBox)
Dim mDwn, I As Long
'    BoxContainer.Cls
    BoxContainer.ScaleMode = vbTwips
    'BoxContainer.AutoRedraw = True
    BoxContainer.BackColor = &HFCF1ED
    BoxContainer.Line (0, 0)-(BoxContainer.ScaleWidth - 10, BoxContainer.ScaleHeight - 20), vbWhite, B '&H6D4016
'    BoxContainer.Line (10, 10)-(BoxContainer.ScaleWidth - 30, BoxContainer.ScaleHeight - 30), &HEAAF6F, B
'    BoxContainer.Line (30, 30)-(BoxContainer.ScaleWidth - 50, BoxContainer.ScaleHeight - 50), &HEAAF6F, B
'    BoxContainer.Line (40, 40)-(BoxContainer.ScaleWidth - 70, BoxContainer.ScaleHeight - 60), &H6D4016, B
    mDwn = 120
    For I = 1 To 13
        BoxContainer.Line (150, mDwn)-(BoxContainer.ScaleWidth - 150, mDwn), &H6D4016
        BoxContainer.Line (150, mDwn + 10)-(BoxContainer.ScaleWidth - 150, mDwn + 10), &HEAAF6F
        mDwn = mDwn + 30
    Next I
'    BoxContainer.Refresh
End Sub

Public Sub HiasForm(ByVal BoxContainer As PictureBox, ByVal MeActive As Form)
Dim maclefttext, mactoptext As Long
Dim obj As Object
MeActive.WindowState = 0
MeActive.BorderStyle = 3
MeActive.BackColor = &HEAAF6F

If MeActive.MDIChild = True Then
   'BoxContainer.Width = MeActive.Width - 120
   'BoxContainer.Height = MeActive.Height - 650
Else
   BoxContainer.Appearance = 0
   BoxContainer.Top = 30
   BoxContainer.Left = 50
   
   BoxContainer.width = MeActive.width - 150
   BoxContainer.Height = MeActive.Height - 650
   BoxContainer.ForeColor = &HFCF1ED
   BoxContainer.FontSize = 16
   BoxContainer.FontBold = False

   Call ColForm(BoxContainer)
   maclefttext = (BoxContainer.ScaleWidth / 2) - ((BoxContainer.TextWidth(MeActive.Tag) / 2))
   BoxContainer.CurrentX = maclefttext '+ (BoxContainer.TextWidth(MeActive.Caption) / 2)
   mactoptext = (BoxContainer.ScaleHeight / 2) - (BoxContainer.TextHeight(MeActive.Tag) / 2)
   BoxContainer.CurrentY = 120 'mactoptext
   BoxContainer.Print MeActive.Tag
End If

'End If
For Each obj In MeActive.Controls
    If TypeOf obj Is TextBox Or TypeOf obj Is MaskEdBox Then
       obj.Appearance = 0
       obj.BorderStyle = 0
       obj.BackColor = &HFCF1ED
       obj.ForeColor = &H6D4016
       obj.FontSize = 8
    End If
    If TypeOf obj Is Label Then
       obj.BackColor = &HFCF1ED
       obj.ForeColor = vbWindowText '&HFFFFFF    '&H6D4016
       obj.FontBold = False
       obj.FontSize = 8
    End If
    If TypeOf obj Is Line Then
       obj.BorderColor = &H6D4016
    End If
    If TypeOf obj Is CommandButton Then
       obj.BackColor = &HEAAF6F
       obj.MaskColor = &HEAAF6F
       obj.FontSize = 8
    End If
    If TypeOf obj Is CheckBox Then
'       obj.BackColor = &HFCF1ED
       obj.BackColor = &HEAAF6F
       obj.ForeColor = &H80000008 '&H6D4016
       obj.FontSize = 8
    End If
    If TypeOf obj Is DataGrid Then
       obj.BackColor = &HFCF1ED
       obj.ForeColor = &H6D4016
       'Dim ob As DataGrid
       obj.Font.Size = 8
       obj.HeadFont.Size = 8
       obj.HeadLines = 1
    End If
    If TypeOf obj Is ComboBox Then
       obj.BackColor = &HFCF1ED
       obj.ForeColor = &H6D4016
       obj.FontSize = 8
    End If
    If TypeOf obj Is DataCombo Then
       obj.BackColor = &HFCF1ED
       obj.ForeColor = &H6D4016
       'Dim ob As DataCombo
       'obj.FontSize = 8
       obj.Font.Size = 8
    End If
Next obj
'BoxContainer.Refresh
End Sub

Public Sub HiasFormCaller(ByVal BoxContainer As PictureBox, ByVal MeActive As Form)
Dim maclefttext, mactoptext As Long
Dim obj As Object
MeActive.WindowState = 0
MeActive.BorderStyle = 3
BoxContainer.Appearance = 0
BoxContainer.Top = 30
BoxContainer.Left = 50
If MeActive.MDIChild = True Then
   BoxContainer.width = MeActive.width - 120
   BoxContainer.Height = MeActive.Height - 650
Else
   BoxContainer.width = MeActive.width - 150
   BoxContainer.Height = MeActive.Height - 650
End If
'End If
BoxContainer.ForeColor = &HFCF1ED
BoxContainer.FontSize = 16
MeActive.BackColor = &HEAAF6F
Call ColForm(BoxContainer)
maclefttext = (BoxContainer.ScaleWidth / 2) - ((BoxContainer.TextWidth(MeActive.Tag) / 2))
BoxContainer.CurrentX = maclefttext '+ (BoxContainer.TextWidth(MeActive.Caption) / 2)
mactoptext = (BoxContainer.ScaleHeight / 2) - (BoxContainer.TextHeight(MeActive.Tag) / 2)
BoxContainer.CurrentY = 120 'mactoptext
BoxContainer.Print MeActive.Tag
For Each obj In MeActive.Controls
    If TypeOf obj Is TextBox Or TypeOf obj Is MaskEdBox Then
       obj.Appearance = 0
       obj.BorderStyle = 0
       obj.BackColor = &HFCF1ED
       obj.ForeColor = &H6D4016
       obj.FontSize = 8
    End If
    If TypeOf obj Is Label Then
       obj.BackColor = &HFCF1ED
       obj.ForeColor = vbWindowText '&HFFFFFF    '&H6D4016
       obj.FontBold = False
       obj.FontSize = 8
    End If
    If TypeOf obj Is Line Then
       obj.BorderColor = &H6D4016
    End If
    If TypeOf obj Is CommandButton Then
       obj.BackColor = &HEAAF6F
       obj.MaskColor = &HEAAF6F
       obj.FontSize = 8
    End If
    If TypeOf obj Is CheckBox Then
       obj.BackColor = &HFCF1ED
       obj.ForeColor = &H6D4016
       obj.FontSize = 8
    End If
    If TypeOf obj Is DataGrid Then
       obj.BackColor = &HFCF1ED
       obj.ForeColor = &H6D4016
       'Dim ob As DataGrid
       obj.Font.Size = 8
       obj.HeadFont.Size = 8
       obj.HeadLines = 1
    End If
    If TypeOf obj Is ComboBox Then
       obj.BackColor = &HFCF1ED
       obj.ForeColor = &H6D4016
       obj.FontSize = 8
    End If
    If TypeOf obj Is DataCombo Then
       obj.BackColor = &HFCF1ED
       obj.ForeColor = &H6D4016
       'Dim ob As DataCombo
       'obj.FontSize = 8
       obj.Font.Size = 8
    End If
Next obj
'BoxContainer.Refresh
End Sub
Public Sub HiasFormManTell(ByVal BoxContainer As PictureBox, ByVal MeActive As Form)
Dim maclefttext, mactoptext As Long
Dim obj As Object
MeActive.WindowState = 0
MeActive.BorderStyle = 3
MeActive.BackColor = &HEAAF6F

For Each obj In MeActive.Controls
    If TypeOf obj Is TextBox Or TypeOf obj Is MaskEdBox Then
       obj.Appearance = 0
       obj.BorderStyle = 0
       obj.BackColor = &HFCF1ED
       obj.ForeColor = &H6D4016
       obj.FontSize = 8
    End If
    If TypeOf obj Is Label Then
       obj.BackColor = &HFCF1ED
       obj.ForeColor = vbWindowText '&HFFFFFF    '&H6D4016
       obj.FontBold = False
       obj.FontSize = 8
'       obj.Alignment = 0
'       obj.AutoSize = False
    End If
    If TypeOf obj Is Line Then
       obj.BorderColor = &HFFFFFF '&H6D4016
    End If
    If TypeOf obj Is CommandButton Then
       obj.BackColor = &HEAAF6F
       obj.MaskColor = &HEAAF6F
       obj.FontSize = 8
    End If
    If TypeOf obj Is CheckBox Then
'       obj.BackColor = &HFCF1ED
       obj.BackColor = &HEAAF6F
       obj.ForeColor = &H80000008 '&H6D4016
       obj.FontSize = 8
    End If
    If TypeOf obj Is DataGrid Then
       obj.BackColor = &HFCF1ED
       obj.ForeColor = &H6D4016
       'Dim ob As DataGrid
       obj.Font.Size = 8
       obj.HeadFont.Size = 8
       obj.HeadLines = 1
    End If
    If TypeOf obj Is ComboBox Then
       obj.BackColor = &HFCF1ED
       obj.ForeColor = &H6D4016
       obj.FontSize = 8
    End If
    If TypeOf obj Is DataCombo Then
       obj.BackColor = &HFCF1ED
       obj.ForeColor = &H6D4016
       'Dim ob As DataCombo
       'obj.FontSize = 8
       obj.Font.Size = 8
    End If
Next obj
CenterForm BoxContainer, MeActive
'BoxContainer.Refresh
End Sub
Public Sub CenterForm(ByVal PanelForm As Object, frm As Form, Optional ByVal Tipical As Boolean = False)
On Error Resume Next
If Tipical = False Then
   'PanelForm.Cls
   'PanelForm.AutoRedraw = True
   'PanelForm.BorderStyle = 0
   'PanelForm.BackColor = &HEAAF6F
   If frm.BorderStyle <> 1 Then
       Select Case frm.WindowState
              Case 0:
                   frm.Top = (MainMenu.ScaleHeight / 2) - ((frm.Height / 2) - 200)
                   frm.Left = (MainMenu.ScaleWidth / 2) - (frm.width / 2)
                   'frm.Width = MainMenu.ScaleWidth
                   'PanelForm.Left = (Frm.ScaleWidth / 2) - (PanelForm.Width / 2) - 40
                   'frm.Height = MainMenu.ScaleHeight - 50
                   'PanelForm.Top = (Frm.ScaleHeight / 2) - ((PanelForm.Height / 2) + 50)
              Case 2:
                   'PanelForm.Top = (Frm.ScaleHeight / 2) - ((PanelForm.Height / 2) + 50)
                   'PanelForm.Left = (Frm.ScaleWidth / 2) - (PanelForm.Width / 2) - 40
    '          Case Else
    '               frm.Top = (MainMenu.ScaleHeight / 2) - ((frm.Height / 2 - MainMenu.CoolBar1.Height + MainMenu.CoolBar2.Height) - 200)
    '               frm.Left = (MainMenu.ScaleWidth / 2) - (frm.Width / 2)
    '               frm.Width = MainMenu.ScaleWidth
    '               frm.Height = MainMenu.ScaleHeight - 50
       End Select
   Else
       frm.Top = (MainMenu.ScaleHeight / 2) - ((frm.Height / 2) - 200)
       frm.Left = (MainMenu.ScaleWidth / 2) - (frm.width / 2)
       'PanelForm.Left = (Frm.ScaleWidth / 2) - (PanelForm.Width / 2) - 40
       'PanelForm.Top = (Frm.ScaleHeight / 2) - ((PanelForm.Height / 2) + 50)
   End If
   'PanelForm.Line (1, 1)-(PanelForm.ScaleWidth - 20, PanelForm.ScaleHeight - 20), vbWhite, B
'   PanelForm.Refresh
Else
   frm.Top = (MainMenu.ScaleHeight / 2) - ((frm.Height / 2) - 200)
   frm.Left = (MainMenu.ScaleWidth / 2) - (frm.width / 2)
   'PanelForm.Left = (Frm.ScaleWidth / 2) - (PanelForm.Width / 2) - 40
   'PanelForm.Top = (Frm.ScaleHeight / 2) - ((PanelForm.Height / 2) + 50)
End If
Err.Clear
End Sub

Public Function PeriodeBerjalan() As Boolean
On Error Resume Next
Dim rcPer As New DBQuick
rcPer.DBOpen " SELECT GlFile, BeginDate, EndDate, Periode,Closed FROM         SettingPeriod WHERE     (Closed = 0) ORDER BY GlFile", CNN, lckLockReadOnly
With rcPer.DBRecordset
     If .Recordcount <> 0 Then
        mVarPeriode = IIf(Not IsNull(.Fields("Periode")), .Fields("Periode"), 1)
        dDateBegin = Format(IIf(Not IsNull(.Fields("BeginDate")), .Fields("BeginDate"), Date), "dd/mm/yyyy")
        dDateEnd = Format(IIf(Not IsNull(.Fields("EndDate")), .Fields("EndDate"), Date), "dd/mm/yyyy")
        PeriodeBerjalan = True
        TahunFiskalYear = CDbl(IIf(Not IsNull(.Fields(0)), Left(.Fields(0), 4), Format(Year(Date), "0###")))
     Else
        PeriodeBerjalan = False
        TahunFiskalYear = Format(Year(Date), "0###")
     End If
End With
rcPer.CloseDB
End Function

'Public Function CariAkun(ByVal NamaForm As String, ByVal NamaValue As String) As String
'Dim rcAkun As New DBQuick
'rcAkun.DBOpen "SELECT [Daftar Configurasi].NoAccount FROM [Daftar Configurasi] INNER JOIN                       GLAccount ON [Daftar Configurasi].NoAccount = GLAccount.NoAccount WHERE     ([Daftar Configurasi].[Nama Form] = N'" & NamaForm & "') AND ([Daftar Configurasi].[Value Data] = N'" & NamaValue & "') GROUP BY [Daftar Configurasi].NoAccount", Cnn, lckLockReadOnly
'With rcAkun.DBRecordset
'     If .Recordcount <> 0 Then
'        CariAkun = IIf(Not IsNull(.Fields(0)), .Fields(0), "")
'     End If
'End With
'End Function

Public Function PeriodeFilter() As Integer
Select Case mVarPeriode
       Case 1: PeriodeFilter = 12
       Case 2: PeriodeFilter = 1
       Case 3: PeriodeFilter = 2
       Case 4: PeriodeFilter = 3
       Case 5: PeriodeFilter = 4
       Case 6: PeriodeFilter = 5
       Case 7: PeriodeFilter = 6
       Case 8: PeriodeFilter = 7
       Case 9: PeriodeFilter = 8
       Case 10: PeriodeFilter = 9
       Case 11: PeriodeFilter = 10
       Case 12: PeriodeFilter = 1
End Select
End Function

Public Function TotalKas(ByVal NoBankID As String) As Variant
Dim RcKas As New DBQuick
Dim strSQL As String

If NoBankID = "" Then NoBankID = "xxxx"
'RcKas.DBOpen "SELECT  ABS(SUM(ISNULL([Tabel Pembantu].CurrentDR" & PeriodeFilter & ", 0) + [Detail Journal].Debet)  - SUM(ISNULL([Tabel Pembantu].CurrentCR" & PeriodeFilter & ", 0) + [Detail Journal].Credit)) AS Saldo FROM         [Table Journal] INNER JOIN [Detail Journal] ON [Table Journal].JournalID = [Detail Journal].JournalID INNER JOIN GLAccount ON [Detail Journal].NoAccount = GLAccount.NoAccount LEFT OUTER JOIN [Tabel Pembantu] ON GLAccount.NoAccount = [Tabel Pembantu].NoAccount WHERE     ([Table Journal].Periode = " & mVarPeriode & ") AND (GLAccount.NoAccount =N'" & NoBankID & "') AND (GLAccount.[Group] = N'Detail List Account') GROUP BY GLAccount.NoAccount, GLAccount.AccountName", CNN, lckLockReadOnly

strSQL = "SELECT ISNULL([tabel pembantu].CurrentDR" & PeriodeFilter & ", 0) AS saldo_debet, ISNULL([tabel pembantu].CurrentCR" & PeriodeFilter & ", 0) AS saldo_kredit, " & _
        " SUM(ISNULL([Detail Journal].Debet, 0)) AS debet, SUM(ISNULL([Detail Journal].Credit, 0)) AS kredit " & _
        " FROM [table journal] INNER JOIN [Detail Journal] ON [table journal].JournalID = [Detail Journal].JournalID " & _
        " INNER JOIN [tabel pembantu] ON [Detail Journal].NoAccount = [tabel pembantu].NoAccount " & _
        " WHERE ([table journal].Periode = " & mVarPeriode & ") AND ([Detail Journal].NoAccount = N'" & NoBankID & "') " & _
        " GROUP BY ISNULL([tabel pembantu].CurrentDR" & PeriodeFilter & ", 0), ISNULL([tabel pembantu].CurrentCR" & PeriodeFilter & ", 0)"
        
RcKas.DBOpen strSQL, CNN, lckLockReadOnly
'Debug.Print RcKas.DBRecordset.Source
With RcKas
     If .Recordcount <> 0 Then
        TotalKas = (.DBRecordset.Fields(0).Value + .DBRecordset.Fields(2).Value) - (.DBRecordset.Fields(1).Value + .DBRecordset.Fields(3).Value)
     Else
        TotalKas = 0
     End If
End With
End Function
'
'Public Function CheckEmptyGrid(ByVal RecName As Recordset, ByVal ListFieldByName As String) As Boolean
'Dim rcc As Recordset
'Dim Avdata As Variant
'Dim I As Integer
'Set rcc = RecName.Clone(adLockReadOnly)
'If rcc.Recordcount <> 0 Then
'   Avdata = rcc.Getrows(rcc.Recordcount, adBookmarkFirst, ListFieldByName)
'   For I = 0 To UBound(Avdata, 2)
'       If IsNull(Avdata(0, I)) = True Then CheckEmptyGrid = True: Exit For
'       If Avdata(0, I) = "" Then CheckEmptyGrid = True: Exit For
'       If Avdata(0, I) = "-" Then CheckEmptyGrid = True: Exit For
'       If Avdata(0, I) = 0 Then CheckEmptyGrid = True: Exit For
'   Next I
'End If
'Set Avdata = Nothing
'End Function

Public Function StayOnTop(Form As Form)
Dim lFlags As Long
Dim lStay As Long
lFlags = SWP_NOSIZE Or SWP_NOMOVE
lStay = SetWindowPos(Form.hwnd, HWND_TOPMOST, 0, 0, 0, 0, lFlags)
End Function

Public Function ReleaseTop(Form As Form)
Call SetWindowPos(Form.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End Function

Public Function CariNoAccount(ByVal Params As String) As String
Dim RcAkum As New DBQuick
RcAkum.DBOpen "SELECT     NoAccount, AccountName FROM         GLAccount WHERE     (Type = N'" & Params & "')", CNN, lckLockReadOnly
With RcAkum.DBRecordset
     If .Recordcount <> 0 Then
        CariNoAccount = IIf(Not IsNull(.Fields(0)), .Fields(0), "")
     End If
End With
End Function

Public Function CariTypeJournal(ByVal Params As Long) As String
Dim RcAkum As New DBQuick
RcAkum.DBOpen "SELECT     GLAccount.NoAccount FROM         AccType INNER JOIN                       GLAccount ON AccType.Tipe = GLAccount.Type WHERE     (GLAccount.[Group] = N'Detail List Account') AND (AccType.ID = " & Params & ")", CNN, lckLockReadOnly
With RcAkum.DBRecordset
     If .Recordcount <> 0 Then
        CariTypeJournal = IIf(Not IsNull(.Fields(0)), .Fields(0), "")
     End If
End With
End Function

'Public Sub UserLogin(ByVal Username As String, Optional ByVal Password As String)
'On Error Resume Next
'Dim Rc As New DBQuick
'Dim Avdata As Variant
'Dim I As Integer
'Dim strSQL As String
'
'''Rc.DBOpen "SELECT [Detail User Table].[Form List], [Detail User Table].Edit, [Detail User Table].Tambah, [Detail User Table].Hapus, [Detail User Table].Laporan,                        [Detail User Table].[Print],[User Table].[User ID] FROM         [Detail User Table] INNER JOIN                       [User Table] ON [Detail User Table].[User ID] = [User Table].[User ID] WHERE     ([User Table].[User Name] Like N'%" & Username & "%') ", Cnn, lckLockReadOnly
'strSQL = "SELECT [Form Table].[Form List], [Detail User Table].Edit, [Detail User Table].Tambah, [Detail User Table].Hapus, [Detail User Table].Laporan,  [Detail User Table].[Print], [User Table].[User ID] FROM         [Detail User Table] INNER JOIN [User Table] ON [Detail User Table].[User ID] = [User Table].[User ID] INNER JOIN    [Form Table] ON [Detail User Table].Idx = [Form Table].Idx WHERE     ([User Table].[User Name] = N'" & Username & "') ORDER BY [Form Table].[Form List]"
'
''Debug.Print strSQL
'Rc.DBOpen strSQL, CNN, lckLockReadOnly
'
'With Rc.DBRecordset
'     If .Recordcount <> 0 Then
''        CloseMenuAll
'        mVarIDUser = ""
'        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
'        For I = 0 To UBound(Avdata, 2)
'            mVarIDUser = Avdata(6, I)
'            Select Case Avdata(0, I)
'                'Master Data
'                Case "Bank Partner": MainMenu.mnBankPartner.Enabled = CBool(Avdata(4, I))
'                Case "Gudang Customer": MainMenu.mnCusGudang.Enabled = CBool(Avdata(4, I))
'                Case "Master Customer": MainMenu.mncus.Enabled = CBool(Avdata(4, I))
'                Case "Master Gudang": MainMenu.MnGudang.Enabled = CBool(Avdata(4, I))
'                Case "Master Item Barang":
'                     MainMenu.mnInvCard.Enabled = CBool(Avdata(4, I))
'                     MainMenu.mnMInventory.Enabled = CBool(Avdata(4, I))
'                Case "Master Karyawan": MainMenu.mnKaryawan.Enabled = CBool(Avdata(4, I))
'                Case "Master Kelompok": MainMenu.mnKelompok.Enabled = CBool(Avdata(4, I))
'                Case "Master Mata Uang": MainMenu.mnCurrency.Enabled = CBool(Avdata(4, I))
'                Case "Tipe Pengiriman": MainMenu.mnTipeBayar.Enabled = CBool(Avdata(4, I))
'                Case "Master Regional": MainMenu.mnRegional.Enabled = CBool(Avdata(4, I))
'                Case "Master Supplier": MainMenu.mnSup.Enabled = CBool(Avdata(4, I))
'                Case "Master Transporter": MainMenu.mnTransport.Enabled = CBool(Avdata(4, I))
'               'PURCHASE
'                  'ORDER
'                Case "Order Pembelian": MainMenu.mnPurchaseOrder.Enabled = CBool(Avdata(4, I))
'                'Case "Penerimaan Barang": MainMenu.mnRn.Enabled = CBool(Avdata(4, I))
'                Case "Order Penjualan": MainMenu.mnPurchaseOrder.Enabled = CBool(Avdata(4, I))
'                'Case "Retur Pembelian": MainMenu.mnReturBeli.Enabled = CBool(Avdata(4, I))
'                Case "Retur Penjualan": MainMenu.mnSalesReturn.Enabled = CBool(Avdata(4, I))
'                Case "Penagihan / Invoice": MainMenu.MnInvSales.Enabled = CBool(Avdata(4, I))
'                'Case "Surat Jalan": MainMenu.mnDn.Enabled = CBool(Avdata(4, I))
'
'                'Menu Akunting
'                Case "Pelunasan Piutang Karyawan": MainMenu.mnBkmPiutang.Enabled = CBool(Avdata(4, I))
'                Case "Penerimaan Tunai Lainnya": MainMenu.mnBkm.Enabled = CBool(Avdata(4, I))
'                Case "Penukaran Setara Kas Ke Kas": MainMenu.mnTukasKas.Enabled = CBool(Avdata(4, I))
'
'                Case "Pengeluaran Piutang Ke Karyawan": MainMenu.mnBkkPiutang.Enabled = CBool(Avdata(4, I))
'                Case "Pengeluaran Tunai Biaya": MainMenu.mnBkk.Enabled = CBool(Avdata(4, I))
'
'                Case "Pelunasan Hutang / Piutang": MainMenu.mnVoucher.Enabled = CBool(Avdata(4, I))
'                Case "Tutup Buku (Periode)": MainMenu.mnClosing.Enabled = CBool(Avdata(4, I))
'
'                Case "Memorial Umum": MainMenu.mnMemoUmum.Enabled = CBool(Avdata(4, I))
'                Case "Memorial Pembelian / Penjualan": MainMenu.mnMemoJualbeli.Enabled = CBool(Avdata(4, I))
'
'                Case "Master Perkiraan": MainMenu.mnPerkiraan.Enabled = CBool(Avdata(4, I))
'                Case "Setup Report Ledger": MainMenu.mnSetupAccount.Enabled = CBool(Avdata(4, I))
'                Case "Seting Periode": MainMenu.mnPeriode.Enabled = CBool(Avdata(4, I))
'
''                Menu Produksi
'                Case "BOM Methode": MainMenu.mnBomMethode.Enabled = CBool(Avdata(4, I))
'                Case "Type Cost": MainMenu.mnTypeCost.Enabled = CBool(Avdata(4, I))
'                Case "Manufacture Stage/Tahap": MainMenu.mnCountPoint.Enabled = CBool(Avdata(4, I))
'                Case "Manufacture Resources Type": MainMenu.mnRscType.Enabled = CBool(Avdata(4, I))
'                Case "Resources": MainMenu.mnRsc.Enabled = CBool(Avdata(4, I))
'                Case "Manufacture Item Kategori": MainMenu.mnItemCategories.Enabled = CBool(Avdata(4, I))
'                Case "Tipe Deskripsi": MainMenu.mnTipeDes.Enabled = CBool(Avdata(4, I))
'                Case "Scheduling Calendar": MainMenu.mnCalendar.Enabled = CBool(Avdata(4, I))
'                Case "Manufacture Work Center": MainMenu.mnManWC.Enabled = CBool(Avdata(4, I))
'                Case "Item Reference": MainMenu.mnItemReference.Enabled = CBool(Avdata(4, I))
'                Case "Manufacture Descriptor": MainMenu.mnDescrip.Enabled = CBool(Avdata(4, I))
'                'Case "Mutasi Gudang": MainMenu.mnMutasi.Enabled = CBool(Avdata(4, I))
'                'Case "Inventory Adjustment": MainMenu.mnInvAdj.Enabled = CBool(Avdata(4, I))
'                Case "Descriptor Reference": MainMenu.mnDescripRef.Enabled = CBool(Avdata(4, I))
'                Case "Lot Sizing": MainMenu.mnLot.Enabled = CBool(Avdata(4, I))
'                Case "Master Schedule": MainMenu.mnSchedule.Enabled = CBool(Avdata(4, I))
'                Case "Job Costing": MainMenu.mnJobCosting.Enabled = CBool(Avdata(4, I))
'                Case "Manufacture Order": MainMenu.mnManOrder.Enabled = CBool(Avdata(4, I))
'                Case "Bill Of Material": MainMenu.mnBomBom.Enabled = CBool(Avdata(4, I))
'
'            End Select
'        Next I
'
'     Else
''        CloseMenuAll
'     End If
'End With
''MainMenu.mnInventory.Visible = False
''IsEnabledLogin = CBool(Avdata(4, i))
'End Sub

Public Sub CloseMenuAll()
'On Error Resume Next
'Master Menu
MainMenu.mnBankPartner.Visible = VisibledMenu(MainMenu.mnMaster, MainMenu.mnBankPartner)
MainMenu.mnCusGudang.Visible = VisibledMenu(MainMenu.mnMaster, MainMenu.mnCusGudang)
MainMenu.mncus.Visible = VisibledMenu(MainMenu.mnMaster, MainMenu.mncus)
MainMenu.MnGudang.Visible = VisibledMenu(MainMenu.mnMaster, MainMenu.MnGudang)
MainMenu.mnInvCard.Visible = VisibledMenu(MainMenu.mnMaster, MainMenu.mnInvCard)
MainMenu.mnKaryawan.Visible = VisibledMenu(MainMenu.mnMaster, MainMenu.mnKaryawan)
MainMenu.mnKelompok.Visible = VisibledMenu(MainMenu.mnMaster, MainMenu.mnKelompok)
MainMenu.mnCurrency.Visible = VisibledMenu(MainMenu.mnMaster, MainMenu.mnCurrency)
MainMenu.mnTipeBayar.Visible = VisibledMenu(MainMenu.mnMaster, MainMenu.mnTipeBayar)
MainMenu.mnRegional.Visible = VisibledMenu(MainMenu.mnMaster, MainMenu.mnRegional)
MainMenu.mnSup.Visible = VisibledMenu(MainMenu.mnMaster, MainMenu.mnSup)
MainMenu.mnTransport.Visible = VisibledMenu(MainMenu.mnMaster, MainMenu.mnTransport)
MainMenu.mnMaster = HiddenMenu(MainMenu.mnMaster, MainMenu.mnBankPartner Or MainMenu.mnCusGudang Or MainMenu.mncus Or _
                               MainMenu.MnGudang Or MainMenu.mnInvCard Or MainMenu.mnKaryawan Or MainMenu.mnKelompok Or MainMenu.mnCurrency Or MainMenu.mnTipeBayar Or _
                               MainMenu.mnRegional Or MainMenu.mnSup Or MainMenu.mnTransport)
'Distribusi Menu
MainMenu.mnPurchase.Visible = VisibledMenu(MainMenu.mnPurchase, MainMenu.mnPurchaseOrder)
'MainMenu.mnPurchase.Visible = VisibledMenu(MainMenu.mnPurchase, MainMenu.mnRn) 'ada perubahan mas enda lama penerimaan
MainMenu.mnPurchase.Visible = VisibledMenu(MainMenu.mnPurchase, MainMenu.mnPermintaanBrg)
MainMenu.mnMarketing.Visible = VisibledMenu(MainMenu.mnMarketing, MainMenu.mnPurchaseOrder)
MainMenu.mnMarketing.Visible = VisibledMenu(MainMenu.mnMarketing, MainMenu.MnInvSales)
'MainMenu.mnMarketing.Visible = VisibledMenu(MainMenu.mnMarketing, MainMenu.mnDn) 'ada perubahan mas enda dlivery order
'MainMenu.mnPurchase.Visible = VisibledMenu(MainMenu.mnPurchase, MainMenu.mnReturBeli) 'ada perubahan
MainMenu.mnMarketing.Visible = VisibledMenu(MainMenu.mnMarketing, MainMenu.mnSalesReturn)
'MainMenu.mnTrans.Visible = HiddenMenu(MainMenu.mnTrans, MainMenu.mnPurchaseOrder Or MainMenu.mnRn Or MainMenu.mnPurchaseOrder Or MainMenu.MnInvSales Or MainMenu.mnDn Or MainMenu.mnReturBeli Or MainMenu.mnSalesReturn)

'Akunting
MainMenu.mnMasAkun.Visible = VisibledMenu(MainMenu.mnMasAkun, MainMenu.mnPerkiraan)
MainMenu.mnKOnfig.Visible = VisibledMenu(MainMenu.mnKOnfig, MainMenu.mnPeriode)
MainMenu.mnKOnfig.Visible = VisibledMenu(MainMenu.mnKOnfig, MainMenu.mnSetupAccount)
MainMenu.mnMasAkun.Visible = HiddenMenu(MainMenu.mnMasAkun, MainMenu.mnPerkiraan.Enabled Or MainMenu.mnPeriode.Enabled Or MainMenu.mnSetupAccount.Enabled)

MainMenu.mnKass.Visible = VisibledMenu(MainMenu.mnKass, MainMenu.mnBkmPiutang)
MainMenu.mnKass.Visible = VisibledMenu(MainMenu.mnKass, MainMenu.mnBkm)
MainMenu.mnKass.Visible = VisibledMenu(MainMenu.mnKass, MainMenu.mnTukasKas)
MainMenu.mnKass.Visible = HiddenMenu(MainMenu.mnKass, MainMenu.mnBkmPiutang Or MainMenu.mnBkm Or MainMenu.mnTukasKas)

MainMenu.mnBKas.Visible = VisibledMenu(MainMenu.mnBKas, MainMenu.mnBkkPiutang)
MainMenu.mnBKas.Visible = VisibledMenu(MainMenu.mnBKas, MainMenu.mnBkk)
MainMenu.mnBKas.Visible = HiddenMenu(MainMenu.mnBKas, MainMenu.mnBkkPiutang Or MainMenu.mnBkk)

MainMenu.mnHutangPiutang.Visible = VisibledMenu(MainMenu.mnHutangPiutang, MainMenu.mnVoucher)
MainMenu.mnHutangPiutang.Visible = HiddenMenu(MainMenu.mnHutangPiutang, MainMenu.mnVoucher)

MainMenu.mnClosed.Visible = VisibledMenu(MainMenu.mnClosed, MainMenu.mnClosing)
MainMenu.mnClosed.Visible = HiddenMenu(MainMenu.mnClosed, MainMenu.mnClosing)

MainMenu.mnMemorial.Visible = VisibledMenu(MainMenu.mnMemorial, MainMenu.mnMemoUmum)
MainMenu.mnMemorial.Visible = VisibledMenu(MainMenu.mnMemorial, MainMenu.mnMemoJualbeli)
MainMenu.mnMemorial.Visible = HiddenMenu(MainMenu.mnMemorial, MainMenu.mnMemoUmum Or MainMenu.mnMemoJualbeli)

MainMenu.mnAkun.Visible = HiddenMenu(MainMenu.mnAkun, MainMenu.mnMasAkun Or MainMenu.mnKass Or MainMenu.mnBKas Or MainMenu.mnHutangPiutang Or MainMenu.mnClosed Or MainMenu.mnMemorial)

'Produksi
MainMenu.mnEntupProduksi.Visible = VisibledMenu(MainMenu.mnEntupProduksi, MainMenu.mnTypeCost)
MainMenu.mnEntupProduksi.Visible = VisibledMenu(MainMenu.mnEntupProduksi, MainMenu.mnBomMethode)
MainMenu.mnEntupProduksi.Visible = VisibledMenu(MainMenu.mnEntupProduksi, MainMenu.mnCountPoint)
MainMenu.mnEntupProduksi.Visible = VisibledMenu(MainMenu.mnEntupProduksi, MainMenu.mnRscType)
MainMenu.mnEntupProduksi.Visible = VisibledMenu(MainMenu.mnEntupProduksi, MainMenu.mnRsc)
MainMenu.mnEntupProduksi.Visible = VisibledMenu(MainMenu.mnEntupProduksi, MainMenu.mnItemCategories)
MainMenu.mnEntupProduksi.Visible = VisibledMenu(MainMenu.mnEntupProduksi, MainMenu.mnTipeDes)
MainMenu.mnEntupProduksi.Visible = VisibledMenu(MainMenu.mnEntupProduksi, MainMenu.mnCalendar)
MainMenu.mnEntupProduksi.Visible = VisibledMenu(MainMenu.mnEntupProduksi, MainMenu.mnManWC)
MainMenu.mnEntupProduksi.Visible = VisibledMenu(MainMenu.mnEntupProduksi, MainMenu.mnItemReference)
MainMenu.mnEntupProduksi.Visible = VisibledMenu(MainMenu.mnEntupProduksi, MainMenu.mnDescripRef)
MainMenu.mnEntupProduksi.Visible = VisibledMenu(MainMenu.mnEntupProduksi, MainMenu.mnDescrip)
MainMenu.mnEntupProduksi.Visible = VisibledMenu(MainMenu.mnEntupProduksi, MainMenu.mnLot)
MainMenu.mnEntupProduksi.Visible = HiddenMenu(MainMenu.mnEntupProduksi, MainMenu.mnBomMethode Or MainMenu.mnCountPoint Or MainMenu.mnRscType Or MainMenu.mnRsc Or _
                                   MainMenu.mnItemCategories Or MainMenu.mnTipeDes Or MainMenu.mnCalendar Or MainMenu.mnManWC Or MainMenu.mnItemReference Or MainMenu.mnDescripRef Or MainMenu.mnDescrip Or _
                                   MainMenu.mnLot)

'MainMenu.mnMenuPersediaan.Visible = VisibledMenu(MainMenu.mnMenuPersediaan, MainMenu.mnMutasi)  ada perubahan enda
'MainMenu.mnMenuPersediaan.Visible = VisibledMenu(MainMenu.mnMenuPersediaan, MainMenu.mnInvAdj) enda
'MainMenu.mnMenuPersediaan = HiddenMenu(MainMenu.mnMenuPersediaan, MainMenu.mnMutasi Or MainMenu.mnInvAdj)

MainMenu.mnMenuProduksi.Visible = VisibledMenu(MainMenu.mnMenuProduksi, MainMenu.mnSchedule)
MainMenu.mnMenuProduksi.Visible = VisibledMenu(MainMenu.mnMenuProduksi, MainMenu.mnJobCosting)
MainMenu.mnMenuProduksi.Visible = VisibledMenu(MainMenu.mnMenuProduksi, MainMenu.mnManOrder)
MainMenu.mnMenuProduksi.Visible = VisibledMenu(MainMenu.mnMenuProduksi, MainMenu.mnBomBom)
MainMenu.mnMenuProduksi.Visible = HiddenMenu(MainMenu.mnMenuProduksi, MainMenu.mnSchedule Or MainMenu.mnJobCosting Or MainMenu.mnManOrder Or MainMenu.mnBomBom)
MainMenu.mnInventory.Visible = HiddenMenu(MainMenu.mnInventory, MainMenu.mnEntupProduksi Or MainMenu.mnMenuPersediaan Or MainMenu.mnMenuProduksi)
End Sub

Private Function VisibledMenu(ByVal ParentMenu As Menu, ByVal ActMenu As Menu) As Boolean
On Error GoTo Hell
    ActMenu.Visible = ActMenu.Enabled
    ParentMenu.Visible = ActMenu.Enabled
    VisibledMenu = ParentMenu.Visible
Exit Function
Hell:
    VisibledMenu = True
    Err.Clear
End Function

Private Function HiddenMenu(ByVal ParentMenu As Menu, ByVal Tipical As Boolean) As Boolean
On Error GoTo Hell
    ParentMenu.Enabled = Tipical
    ParentMenu.Visible = ParentMenu.Enabled
    HiddenMenu = ParentMenu.Visible
Exit Function
Hell:
    HiddenMenu = ParentMenu.Visible
    Err.Clear
End Function

Public Function MoveForm(TheForm)
    Dim ret
    ReleaseCapture
    SendMessage TheForm, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Function

Public Function Konversi(ByVal nNilai As Currency) As String
Dim Grade As Variant
Dim strTerbilang As String
Dim strPart As String
Dim iGrade As Byte
Grade = Array(" milyar ", " juta ", " ribu ", "")
strTerbilang = ""
If Len(CStr(nNilai)) > 12 Then
   strTerbilang = "Melebihi batas maksimum"
Else: strPart = Format(nNilai, String(12, "0"))
   For iGrade = 1 To 4
      If iGrade = 3 And Val(Mid(strPart, (iGrade - 1) * 3 + 1, 3)) = 1 Then
         'Nilai nominal seribu
         strTerbilang = "seribu"
      ElseIf Val(Mid(strPart, (iGrade - 1) * 3 + 1, 3)) > 0 Then
         strTerbilang = strTerbilang & Generate(Mid(strPart, (iGrade - 1) * 3 + 1, 3), iGrade)
         strTerbilang = strTerbilang & Grade(iGrade - 1)
      End If
   Next iGrade
End If
Konversi = strTerbilang & " Rupiah"
End Function

Private Function Generate(ByVal strPart As String, ByVal iGrade As Byte) As String
Dim Angka1 As Variant
Dim Angka2 As Variant
Dim I As Integer
Dim strHasil As String
Dim nTemp As Byte

Angka1 = Array("Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan")
Angka2 = Array(" ratus ", " puluh ", "")

For I = 1 To 3
   nTemp = Val(Mid(strPart, I, 1))
   If nTemp = 1 Then
      If I = 1 Then
         strHasil = "seratus "
      ElseIf I = 2 Then
         I = I + 1
         nTemp = Val(Mid(strPart, I, 1))
         If nTemp = 0 Then
            strHasil = strHasil & "sepuluh"
         ElseIf nTemp = 1 Then
            strHasil = strHasil & "sebelas"
         Else
            strHasil = strHasil & Angka1(nTemp - 1) & " belas"
         End If
      ElseIf Val(strPart) = 1 And iGrade = 3 Then
         strHasil = strHasil & "se"
      Else
         strHasil = strHasil & "satu"
      End If
   ElseIf nTemp <> 0 Then
      strHasil = strHasil + Angka1(nTemp - 1) + Angka2(I - 1)
   End If
Next I
Generate = strHasil
End Function

Public Function PrintPermitted(ByVal ReportName As String) As Boolean
Dim Rc As New DBQuick
Rc.DBOpen "SELECT     [Report Permit].Laporan FROM         [Report Permit] INNER JOIN [Report Modules] ON [Report Permit].IDReport = [Report Modules].IDReport WHERE     ([Report Permit].[User ID] = " & mVarIDUser & ") AND ([Report Modules].FileNameReport = N'" & ReportName & "')", CNN, lckLockReadOnly
'MsgBox "SELECT     [Report Permit].Laporan FROM         [Report Permit] INNER JOIN [Report Modules] ON [Report Permit].IDReport = [Report Modules].IDReport WHERE     ([Report Permit].[User ID] = " & mVarIDUser & ") AND ([Report Modules].FileNameReport = N'" & ReportName & "')"
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        PrintPermitted = IIf(Not IsNull(.Fields(0)), CBool(.Fields(0)), False)
     Else
        PrintPermitted = False
     End If
     Screen.MousePointer = vbDefault
     If PrintPermitted = False Then MessageBox "Anda tidak mempunyai hak akses atas report [" & UCase(ReportName) & "]" & vbCrLf & "Silahkan anda untuk menghubungi administrator.", "Peringatan", msgOkOnly, msgExclamation
End With
End Function

Public Sub CopyRecordsetByQuery(ByVal SQLQuery As String, ByVal GetTableName As String)
Dim Rc As New DBQuick
Dim Fld As Field
Dim strListField, strListFieldValue As String
Dim I As Integer
Rc.DBOpen SQLQuery, CNN, lckLockReadOnly
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        For Each Fld In .Fields
            strListField = strListField & "[" & Fld.Name & "]" & ","
            Select Case Fld.Type
                   Case 202: strListFieldValue = strListFieldValue & "N'" & Fld.Value & "',"
                   Case 6: strListFieldValue = strListFieldValue & CDbl(Fld.Value) & ","
            End Select
        Next
        SendDataToServer " INSERT TO " & GetTableName & _
                         " (" & strListField & ")" & _
                         " Values (" & strListFieldValue & ")"
     End If
End With
End Sub

Public Function SetDate(ByVal InputDate As Variant) As Date
Dim tgl As Date
MsgBox CDate(Format(InputDate, "dd/mm/yyyy"))
If IsDate(InputDate) = False Then InputDate = DateSerial(Year(CDate(Format(InputDate, "dd/mm/yyyy"))), Month(CDate(Format(InputDate, "dd/mm/yyyy"))), Day(CDate(Format(InputDate, "dd/mm/yyyy"))))
SetDate = DateSerial(Year(InputDate), Month(InputDate), Day(InputDate))
End Function

Public Function IsConfigReady() As Boolean
Dim Rc As New DBQuick
Dim I As Integer
Dim Avdata As Variant
Dim strSQL As String

strSQL = "SELECT AccType.Tipe, GLAccount.NoAccount " & _
        " FROM GLAccount RIGHT OUTER JOIN AccType ON GLAccount.ID = AccType.ID " & _
        " Where (GLAccount.ID Is Null) And (AccType.status = 1)"
Rc.DBOpen strSQL, CNN, lckLockReadOnly

'Rc.DBOpen "SELECT GLAccount.NoAccount, AccType.Tipe FROM GLAccount RIGHT OUTER JOIN AccType ON GLAccount.Type = AccType.Tipe " & _
        " WHERE (dbo.AccType.status = 1) AND (GLAccount.NoAccount IS NULL) ORDER BY GLAccount.NoAccount", CNN, lckLockReadOnly
'Debug.Print Rc.DBRecordset.Source

With Rc
     If .Recordcount <> 0 Then
        IsConfigReady = False
     Else
        IsConfigReady = True
     End If
End With
End Function

Public Function BukaComboDetailQueri(Combonya As DataCombo, _
                                     Tsql As String) As ADODB.Recordset
  Dim rcc As New ADODB.Recordset
  Set rcc = New ADODB.Recordset
  rcc.CursorLocation = adUseClient
  rcc.Open Tsql, CNN, adOpenKeyset, adLockOptimistic
  Combonya.ListField = rcc.Fields("Alias Report").Name

  If rcc.Fields.Count > 1 Then Combonya.BoundColumn = rcc.Fields("NoIdx").Name
  Set BukaComboDetailQueri = rcc
  Set Combonya.RowSource = rcc
End Function

Public Function BukaComboDetailQueriGroup(Combonya As DataCombo, _
                                     Tsql As String) As ADODB.Recordset
  Dim rcc As New ADODB.Recordset
  Set rcc = New ADODB.Recordset
  rcc.CursorLocation = adUseClient
  rcc.Open Tsql, CNN, adOpenKeyset, adLockOptimistic
  Combonya.ListField = rcc.Fields("group name").Name

  If rcc.Fields.Count > 1 Then Combonya.BoundColumn = rcc.Fields("NoIdx").Name
  Set BukaComboDetailQueriGroup = rcc
  Set Combonya.RowSource = rcc
End Function
Public Function EmtpyToNum(ByVal vValue As Variant) As Variant
EmtpyToNum = IIf(Not IsEmpty(vValue), vValue, 0)
End Function

Public Function FDate(vType As ModeCheckDate, vDataCheck As String, vMaskObject As MaskEdBox)
    Select Case vType
        Case ModeCheckDate.TimeData
            FDate = IIf(Len(vDataCheck) = 0, "NULL", "   '1899-12-30 " & vMaskObject.Text & "'")
        Case ModeCheckDate.DateData
            FDate = IIf(Len(vDataCheck) = 0, "NULL", "   '" & Format(vMaskObject, "yyyy-MM-dd") & "'")
    End Select
End Function
Public Function FDatePicker(vType As ModeCheckDate, vDateObject As Date)
Select Case vType
    Case ModeCheckDate.DateData
        FDatePicker = " '" & Format(vDateObject, "yyyy-MM-dd") & "'"
    Case ModeCheckDate.TimeData
        FDatePicker = "'1899-12-30 " & Format(vDateObject, "hh:mm:ss") & "'"
End Select
End Function
Public Function FDateGrid(vType As ModeCheckDate, vDateObject As Date)
If IsNull(vDateObject) Then
    FDateGrid = "Null"
Else
    Select Case vType
        Case ModeCheckDate.DateData
            FDateGrid = "   '" & Format(vDateObject, "yyyy-MM-dd") & "'"
        Case ModeCheckDate.TimeData
            FDateGrid = "  '1899-12-30 " & Format(vDateObject, "hh:mm:ss") & "'"
    End Select
End If
End Function
Public Function FCombo(vData)
    If Len(vData) = 0 Then
        FCombo = "NULL"
    Else
        FCombo = "'" & Replace(vData, "'", "''") & "'"
    End If
End Function
Public Function FNumText(vData)
    If Len(vData) = 0 Then
        FNumText = "NULL"
    Else
        If IsNull(vData) Then
            FNumText = "NULL"
        Else
            FNumText = "N'" & Replace(vData, "'", "''") & "'"
        End If
    End If
End Function
Public Function FQty(vData)
On Error GoTo xErr
    If Len(vData) = 0 Then
        FQty = "NULL"
    Else
        If IsNull(vData) Then
            FQty = 0
        Else
            'FQty = Replace(Round(CDbl(vData), 2), ",", ".")
            FQty = Replace(vData, ",", ".")
        End If
    End If
Exit Function
xErr:
    FQty = 0
    Err.Clear
End Function
Public Function FQtyUser(vData)
On Error GoTo xErr
    If Len(vData) = 0 Then
        FQtyUser = "NULL"
    Else
        If IsNull(vData) Then
            FQtyUser = 0
        Else
            FQtyUser = Replace(Round(CDbl(vData), 2), ",", ".")
            
        End If
    End If
Exit Function
xErr:
    FQtyUser = 0
    Err.Clear
End Function
Public Function FPrice(vData)
    FPrice = Replace(Round(CDbl(vData)), ",", ".")
End Function
Public Function FText(vData)
If Len(vData) = 0 Then
    FText = "NULL"
Else
    If IsNull(vData) Then
        FText = "NULL"
    Else
        FText = "'" & Replace(vData, "'", "''") & "'"
    End If
End If
End Function
Public Function FGaris(vData)
    FGaris = Replace(vData, "'", "'-'")
End Function

Public Function EncodeStr64(encString As String, ByVal MaxPerLine As Integer) As String
' Return radix64 of string
Dim abOutput()  As Byte
Dim sLast       As String
Dim b(3)        As Byte
Dim j           As Integer
Dim CharCount   As Integer
Dim iIndex      As Long
Dim Umax        As Long
Dim I As Long, nLen As Long, nQuants As Long
EncodeStr64 = ""
nLen = Len(encString)
nQuants = nLen \ 3
iIndex = 0
If MaxPerLine < 10 Then MaxPerLine = 10
Umax = nQuants + 1
Call MakeEncTab
If (nQuants > 0) Then
    ReDim abOutput(nQuants * 4 - 1)
    For I = 0 To nQuants - 1
        For j = 0 To 2
            b(j) = Asc(Mid(encString, (I * 3) + j + 1, 1))
        Next
        Call EncodeQuantumB(b)
        abOutput(iIndex) = b(0)
        abOutput(iIndex + 1) = b(1)
        abOutput(iIndex + 2) = b(2)
        abOutput(iIndex + 3) = b(3)
        CharCount = CharCount + 4
        ' insert CRLF if max char per line is reached
        If CharCount >= MaxPerLine Then
            ReDim Preserve abOutput(UBound(abOutput) + 2)
            CharCount = 0
            abOutput(iIndex + 4) = 13
            abOutput(iIndex + 5) = 10
            iIndex = iIndex + 6
            Else
            iIndex = iIndex + 4
            End If
    Next
    EncodeStr64 = StrConv(abOutput, vbUnicode)
End If
Select Case nLen Mod 3
Case 0
    sLast = ""
Case 1
    b(0) = Asc(Mid(encString, nLen, 1))
    b(1) = 0
    b(2) = 0
    Call EncodeQuantumB(b)
    sLast = StrConv(b(), vbUnicode)
    sLast = Left(sLast, 2) & "=="
Case 2
    b(0) = Asc(Mid(encString, nLen - 1, 1))
    b(1) = Asc(Mid(encString, nLen, 1))
    b(2) = 0
    Call EncodeQuantumB(b)
    sLast = StrConv(b(), vbUnicode)
    sLast = Left(sLast, 3) & "="
End Select
EncodeStr64 = EncodeStr64 & sLast
End Function

Public Function PROBAencodeString(aText As String, ByVal aKey As String, expandKey As Boolean) As String
'encrypt string
Dim I As Double
If expandKey = False Then
    'set 32 byte key
    Call PROBAsetKey(aKey)
    Else
    'set variable key
    Call PROBAsetExpandedKey(aKey)
    End If
For I = 1 To Len(aText)
    PROBAencodeString = PROBAencodeString & Chr(PROBAencodeByte(Asc(Mid(aText, I, 1))))
    If I Mod 100 = 0 Then
        UpdateStatus (I / Len(aText) * 100)
    End If
Next
UpdateStatus (0)
End Function

Private Sub PROBAsetKey(aKey As String)
' INPUT: key string containing 32 bytes
' key 256 bit (32 x 8 bit value)
' 32 x 4 (128) bits for sbox selection
' 32 x 4 (128) bits for sbox startposition selection
Dim I As Byte
Dim Ks(32) As Byte
Dim Key() As Byte
Key() = StrConv(aKey, vbFromUnicode)
'select the 32 sBoxes
For I = 0 To 15
    sBox(I + 1) = Key(I) And 15 'Hi boxes
    sBox(I + 17) = Int(Key(I) / 16)  'Lo boxes
Next
' select initial position of the 32 sboxes
For I = 0 To 15
    sBoxPos(I + 1) = Key(I + 16) And 15 'Hi boxes
    sBoxPos(I + 17) = Int(Key(I + 16) / 16) 'Lo boxes
Next
Call PROBAinitKey
End Sub

Private Sub PROBAsetExpandedKey(aKey As String)
' INPUT: variable lenght key,
' recalculated to a 256 bits key with RC4-style scramble
'
' key string containing 32 bytes
' key 256 bit (32 x 8 bits)
' 32 x 4 (128) bits for sbox selection
' 32 x 4 (128) bits for sbox startposition selection
Dim Ks(255) As Integer
Dim ss As Integer
Dim Ps As Integer
Dim I As Integer
Dim j As Integer
Dim KeyLen As Integer
Dim Key() As Byte
Dim Tmp As Byte
For I = 0 To 255
    Ks(I) = I
Next
'scramble (RC4-style)
KeyLen = Len(aKey)
Key() = StrConv(aKey, vbFromUnicode)
For I = 0 To 255
    j = (j + Ks(I) + Key(I Mod KeyLen)) Mod 255
    Tmp = Ks(I)
    Ks(I) = Ks(j)
    Ks(j) = Tmp
Next
'select sBoxs
For I = 1 To 16
    sBox(I) = Ks(I) And 15 'Hi boxes
    sBox(I + 16) = Int(Ks(I) / 16) 'Lo boxes
Next
'select init position sBoxs
For I = 1 To 16
    sBoxPos(I) = Ks(16 + I) And 15 'Hi boxes
    sBoxPos(I + 16) = Int(Ks(16 + I) / 16) 'Lo boxes
Next
Call PROBAinitKey
End Sub

Private Function PROBAencodeByte(aByte As Byte) As Byte
'encrypt High Nible
sBoxOut(1) = sBoxEncode((Int(aByte / 16)), 1)
sBoxOut(2) = sBoxEncode((sBoxOut(1)), 2)
sBoxOut(3) = sBoxEncode((sBoxOut(2)), 3)
sBoxOut(4) = sBoxEncode((sBoxOut(3)), 4)
sBoxOut(5) = sBoxEncode((sBoxOut(4)), 5)
sBoxOut(6) = sBoxEncode((sBoxOut(5)), 6)
sBoxOut(7) = sBoxEncode((sBoxOut(6)), 7)
sBoxOut(8) = sBoxEncode((sBoxOut(7)), 8)
sBoxOut(9) = sBoxEncode((sBoxOut(8)), 9)
sBoxOut(10) = sBoxEncode((sBoxOut(9)), 10)
sBoxOut(11) = sBoxEncode((sBoxOut(10)), 11)
sBoxOut(12) = sBoxEncode((sBoxOut(11)), 12)
sBoxOut(13) = sBoxEncode((sBoxOut(12)), 13)
sBoxOut(14) = sBoxEncode((sBoxOut(13)), 14)
sBoxOut(15) = sBoxEncode((sBoxOut(14)), 15)
sBoxOut(16) = sBoxEncode((sBoxOut(15)), 16)
'encrypt Low Nible
sBoxOut(17) = sBoxEncode((aByte And 15), 17)
sBoxOut(18) = sBoxEncode((sBoxOut(17)), 18)
sBoxOut(19) = sBoxEncode((sBoxOut(18)), 19)
sBoxOut(20) = sBoxEncode((sBoxOut(19)), 20)
sBoxOut(21) = sBoxEncode((sBoxOut(20)), 21)
sBoxOut(22) = sBoxEncode((sBoxOut(21)), 22)
sBoxOut(23) = sBoxEncode((sBoxOut(22)), 23)
sBoxOut(24) = sBoxEncode((sBoxOut(23)), 24)
sBoxOut(25) = sBoxEncode((sBoxOut(24)), 25)
sBoxOut(26) = sBoxEncode((sBoxOut(25)), 26)
sBoxOut(27) = sBoxEncode((sBoxOut(26)), 27)
sBoxOut(28) = sBoxEncode((sBoxOut(27)), 28)
sBoxOut(29) = sBoxEncode((sBoxOut(28)), 29)
sBoxOut(30) = sBoxEncode((sBoxOut(29)), 30)
sBoxOut(31) = sBoxEncode((sBoxOut(30)), 31)
sBoxOut(32) = sBoxEncode((sBoxOut(31)), 32)
'calculate encrypted byte
PROBAencodeByte = (sBoxOut(16) * 16) + sBoxOut(32)
'advance boxes
Call PROBAturnBoxes
End Function

Public Sub UpdateStatus(ByVal sngPercent As Single)
'With FormDemo.picProgress
If sngPercent > 100 Then sngPercent = 100
'If sngPercent = 0 Then '.Cls: Exit Sub
'.DrawMode = 13
'FormDemo.picProgress.Line (-10, -10)-(sngPercent, .Height + 75), .ForeColor, BF
'.Refresh
'End With
End Sub

' ------------------------------------------------------------
'                      Encryption functions
' ------------------------------------------------------------

Private Sub PROBAinitKey()
Dim I As Byte
Dim j As Byte
'sBox configuration (encode)
sBoxInit(0) = Array(12, 8, 9, 1, 2, 4, 10, 13, 11, 3, 0, 15, 7, 6, 14, 5)
sBoxInit(1) = Array(4, 0, 10, 11, 3, 14, 9, 8, 13, 1, 2, 7, 5, 6, 12, 15)
sBoxInit(2) = Array(0, 12, 15, 2, 4, 3, 9, 13, 1, 10, 8, 11, 14, 5, 7, 6)
sBoxInit(3) = Array(7, 6, 8, 5, 0, 9, 3, 2, 1, 10, 15, 11, 14, 4, 13, 12)
sBoxInit(4) = Array(11, 13, 4, 3, 9, 10, 5, 1, 8, 12, 6, 14, 7, 15, 2, 0)
sBoxInit(5) = Array(14, 3, 4, 1, 0, 10, 5, 11, 2, 15, 6, 8, 12, 13, 9, 7)
sBoxInit(6) = Array(1, 5, 11, 12, 6, 4, 15, 0, 7, 3, 14, 9, 13, 8, 10, 2)
sBoxInit(7) = Array(5, 3, 11, 13, 2, 1, 12, 10, 0, 4, 7, 6, 14, 8, 15, 9)
sBoxInit(8) = Array(11, 3, 10, 5, 1, 14, 12, 13, 15, 2, 7, 8, 6, 0, 9, 4)
sBoxInit(9) = Array(1, 2, 6, 0, 15, 5, 13, 3, 14, 4, 10, 12, 9, 11, 8, 7)
sBoxInit(10) = Array(12, 11, 13, 3, 2, 14, 9, 4, 1, 10, 8, 7, 0, 6, 5, 15)
sBoxInit(11) = Array(3, 10, 4, 5, 0, 9, 6, 8, 7, 11, 12, 13, 2, 15, 14, 1)
sBoxInit(12) = Array(7, 0, 9, 8, 3, 10, 13, 1, 11, 4, 2, 12, 6, 14, 5, 15)
sBoxInit(13) = Array(5, 11, 4, 3, 2, 0, 12, 1, 15, 14, 6, 10, 9, 13, 7, 8)
sBoxInit(14) = Array(3, 2, 1, 10, 11, 9, 15, 4, 5, 14, 13, 0, 6, 7, 12, 8)
sBoxInit(15) = Array(2, 6, 13, 0, 15, 14, 12, 9, 8, 11, 3, 10, 5, 7, 4, 1)
'sBoxInv configuration (decode)
For I = 0 To 15
    For j = 0 To 15
    sBoxInvInit(I, sBoxInit(I)(j)) = j
    Next
Next
'turnover points per sBox (first value (0) not used!)
BoxTurnOver = Array(0, 2, 14, 6, 15, 3, 7, 11, 5, 9, 3, 14, 13, 4, 6, 8, 1)
End Sub

Private Function sBoxEncode(aByte As Byte, aBox As Byte) As Byte
'encrypt nible with given sBox and offset
Dim pos As Byte
pos = aByte + sBoxPos(aBox)
If pos > 15 Then pos = pos - 16
sBoxEncode = sBoxInit(sBox(aBox))(pos)
End Function

Private Sub PROBAturnBoxes()
'advance the sBoxes by turnover or by inter-action
Dim I As Byte
'HI sBoxes, normal turns
Rotate (1)
For I = 1 To 15
If sBoxPos(I) = BoxTurnOver(sBox(I)) Then
    Rotate (I + 1)
    Else
    Exit For
    End If
Next
'LO sBoxes, normal turns
Rotate 17
For I = 17 To 31
If sBoxPos(I) = BoxTurnOver(sBox(I)) Then
    Rotate (I + 1)
    Else
    Exit For
    End If
Next
'output depended turns
If sBoxOut(1) = 0 Then Rotate (26)
If sBoxOut(1) = 0 Then Rotate (23)
If sBoxOut(17) = 0 Then Rotate (14)
If sBoxOut(17) = 0 Then Rotate (8)
If sBoxOut(3) = 0 Then Rotate (21)
If sBoxOut(18) = 0 Then Rotate (6)
If sBoxOut(2) = 0 And sBoxOut(4) = 0 Then Rotate (28)
If sBoxOut(7) = 0 And sBoxOut(12) = 0 Then Rotate (15)
If sBoxOut(20) = 0 And sBoxOut(24) = 0 Then Rotate (7)
If sBoxOut(5) = 0 And sBoxOut(6) = 0 Then Rotate (31)
If sBoxOut(18) = 0 And sBoxOut(20) = 0 Then Rotate (17)
If sBoxOut(6) + sBoxOut(27) = 8 Then Rotate (25)
If sBoxOut(10) + sBoxOut(19) = 8 Then Rotate (5)
If sBoxOut(8) + sBoxOut(21) = 8 Then Rotate (30)
If sBoxOut(7) + sBoxOut(19) = 8 Then Rotate (9)
If sBoxOut(4) + sBoxOut(7) = 8 Then Rotate (10)
If sBoxOut(2) + sBoxOut(19) = 15 Then Rotate (32)
If sBoxOut(3) + sBoxOut(22) = 15 Then Rotate (16)
If sBoxOut(6) + sBoxOut(21) = 15 Then Rotate (11)
If sBoxOut(7) + sBoxOut(19) = 15 Then Rotate (19)

'next lines are for demonstration purposes only
' and will visualize the rotations of the Sboxes
'Dim tmp As String
'If FormDemo.chkSbox.value = 0 Then Exit Sub
'FormDemo.lblpos.Caption = ""
'For i = 1 To 16
'tmp = Trim(Str(sBoxPos(i))) & " "
'If Len(tmp) = 2 Then tmp = "0" & tmp
'FormDemo.lblpos.Caption = FormDemo.lblpos.Caption & tmp
'Next
'FormDemo.lblpos.Caption = FormDemo.lblpos.Caption & "- - "
'For i = 17 To 32
'tmp = Trim(Str(sBoxPos(i))) & " "
'If Len(tmp) = 2 Then tmp = "0" & tmp
'FormDemo.lblpos.Caption = FormDemo.lblpos.Caption & tmp
'Next
'FormDemo.lblpos.Refresh
End Sub

Private Sub Rotate(aPos As Byte)
'advance a sBox position by 1
sBoxPos(aPos) = sBoxPos(aPos) + 1
If sBoxPos(aPos) > 15 Then sBoxPos(aPos) = sBoxPos(aPos) - 16
End Sub

Private Function MakeEncTab()
Dim I As Integer
Dim C As Integer
I = 0
For C = Asc("A") To Asc("Z")
    aEncTab(I) = C
    I = I + 1
Next
For C = Asc("a") To Asc("z")
    aEncTab(I) = C
    I = I + 1
Next
For C = Asc("0") To Asc("9")
    aEncTab(I) = C
    I = I + 1
Next
C = Asc("+")
aEncTab(I) = C
I = I + 1
C = Asc("/")
aEncTab(I) = C
I = I + 1
End Function

Private Sub EncodeQuantumB(b() As Byte)
Dim b0 As Byte, b1 As Byte, b2 As Byte, b3 As Byte
b0 = SHR2(b(0)) And &H3F
b1 = SHL4(b(0) And &H3) Or (SHR4(b(1)) And &HF)
b2 = SHL2(b(1) And &HF) Or (SHR6(b(2)) And &H3)
b3 = b(2) And &H3F
b(0) = aEncTab(b0)
b(1) = aEncTab(b1)
b(2) = aEncTab(b2)
b(3) = aEncTab(b3)
End Sub

Private Function SHR2(ByVal bytValue As Byte) As Byte
SHR2 = bytValue \ &H4
End Function

Private Function SHL4(ByVal bytValue As Byte) As Byte
SHL4 = (bytValue * &H10) And &HFF
End Function

Private Function SHR4(ByVal bytValue As Byte) As Byte
SHR4 = bytValue \ &H10
End Function

Private Function SHR6(ByVal bytValue As Byte) As Byte
SHR6 = bytValue \ &H40
End Function

Private Function SHL2(ByVal bytValue As Byte) As Byte
SHL2 = (bytValue * &H4) And &HFF
End Function


Public Function PROBAdecodeString(aText As String, ByVal aKey As String, expandKey As Boolean) As String
'decrypt string
Dim I As Double
If expandKey = False Then
    'set 32 byte key
    Call PROBAsetKey(aKey)
    Else
    'set variable key
    Call PROBAsetExpandedKey(aKey)
    End If
For I = 1 To Len(aText)
    PROBAdecodeString = PROBAdecodeString & Chr(PROBAdecodeByte(Asc(Mid(aText, I, 1))))
    If I Mod 100 = 0 Then
        UpdateStatus (I / Len(aText) * 100)
    End If
Next
UpdateStatus (0)
End Function



Public Function DecodeStr64(decString As String) As String
' Return string of decoded values from radix64
Dim abDecoded() As Byte
Dim d(3)    As Byte
Dim C       As Integer
Dim di      As Integer
Dim I       As Long
Dim nLen    As Long
Dim iIndex  As Long
Dim Umax    As Long
nLen = Len(decString)
If nLen < 4 Then
    Exit Function
End If
ReDim abDecoded(((nLen \ 4) * 3) - 1)
Umax = nLen
iIndex = 0
di = 0
Call MakeDecTab
For I = 1 To Len(decString)
    C = CByte(Asc(Mid(decString, I, 1)))
    C = aDecTab(C)
    If C >= 0 Then
        d(di) = CByte(C)
        di = di + 1
        If di = 4 Then
            abDecoded(iIndex) = SHL2(d(0)) Or (SHR4(d(1)) And &H3)
            iIndex = iIndex + 1
            abDecoded(iIndex) = SHL4(d(1) And &HF) Or (SHR2(d(2)) And &HF)
            iIndex = iIndex + 1
            abDecoded(iIndex) = SHL6(d(2) And &H3) Or d(3)
            iIndex = iIndex + 1
            If d(3) = 64 Then
                iIndex = iIndex - 1
                abDecoded(iIndex) = 0
            End If
            If d(2) = 64 Then
                iIndex = iIndex - 1
                abDecoded(iIndex) = 0
            End If
            di = 0
        End If
    End If
Next I
DecodeStr64 = StrConv(abDecoded(), vbUnicode)
DecodeStr64 = Left(DecodeStr64, iIndex)
End Function

Private Function MakeDecTab()
Dim T As Integer
Dim C As Integer
For C = 0 To 255
    aDecTab(C) = -1
Next
T = 0
For C = Asc("A") To Asc("Z")
    aDecTab(C) = T
    T = T + 1
Next
For C = Asc("a") To Asc("z")
    aDecTab(C) = T
    T = T + 1
Next
For C = Asc("0") To Asc("9")
    aDecTab(C) = T
    T = T + 1
Next
C = Asc("+")
aDecTab(C) = T
T = T + 1
C = Asc("/")
aDecTab(C) = T
T = T + 1
C = Asc("=")
aDecTab(C) = T
End Function

Private Function SHL6(ByVal bytValue As Byte) As Byte
SHL6 = (bytValue * &H40) And &HFF
End Function

Private Function PROBAdecodeByte(aByte As Byte) As Byte
Dim HI As Byte
Dim LO As Byte
'decrypt High Nible
sBoxOut(16) = Int(aByte / 16)
sBoxOut(15) = sBoxDecode((sBoxOut(16)), 16)
sBoxOut(14) = sBoxDecode((sBoxOut(15)), 15)
sBoxOut(13) = sBoxDecode((sBoxOut(14)), 14)
sBoxOut(12) = sBoxDecode((sBoxOut(13)), 13)
sBoxOut(11) = sBoxDecode((sBoxOut(12)), 12)
sBoxOut(10) = sBoxDecode((sBoxOut(11)), 11)
sBoxOut(9) = sBoxDecode((sBoxOut(10)), 10)
sBoxOut(8) = sBoxDecode((sBoxOut(9)), 9)
sBoxOut(7) = sBoxDecode((sBoxOut(8)), 8)
sBoxOut(6) = sBoxDecode((sBoxOut(7)), 7)
sBoxOut(5) = sBoxDecode((sBoxOut(6)), 6)
sBoxOut(4) = sBoxDecode((sBoxOut(5)), 5)
sBoxOut(3) = sBoxDecode((sBoxOut(4)), 4)
sBoxOut(2) = sBoxDecode((sBoxOut(3)), 3)
sBoxOut(1) = sBoxDecode((sBoxOut(2)), 2)
HI = sBoxDecode((sBoxOut(1)), 1)
'decrypt Low Nible
sBoxOut(32) = aByte And 15
sBoxOut(31) = sBoxDecode((sBoxOut(32)), 32)
sBoxOut(30) = sBoxDecode((sBoxOut(31)), 31)
sBoxOut(29) = sBoxDecode((sBoxOut(30)), 30)
sBoxOut(28) = sBoxDecode((sBoxOut(29)), 29)
sBoxOut(27) = sBoxDecode((sBoxOut(28)), 28)
sBoxOut(26) = sBoxDecode((sBoxOut(27)), 27)
sBoxOut(25) = sBoxDecode((sBoxOut(26)), 26)
sBoxOut(24) = sBoxDecode((sBoxOut(25)), 25)
sBoxOut(23) = sBoxDecode((sBoxOut(24)), 24)
sBoxOut(22) = sBoxDecode((sBoxOut(23)), 23)
sBoxOut(21) = sBoxDecode((sBoxOut(22)), 22)
sBoxOut(20) = sBoxDecode((sBoxOut(21)), 21)
sBoxOut(19) = sBoxDecode((sBoxOut(20)), 20)
sBoxOut(18) = sBoxDecode((sBoxOut(19)), 19)
sBoxOut(17) = sBoxDecode((sBoxOut(18)), 18)
LO = sBoxDecode((sBoxOut(17)), 17)
'calculate decrypted byte
PROBAdecodeByte = (HI * 16) + LO
'advance boxes
Call PROBAturnBoxes
End Function

Private Function sBoxDecode(aByte As Byte, aBox As Byte) As Byte
'decrypt nible with given sBoxInv and offset
Dim I As Integer
I = sBoxInvInit(sBox(aBox), aByte)
I = I - sBoxPos(aBox)
If I < 0 Then I = I + 16
sBoxDecode = I
End Function



Sub OpenTable(ByRef vRSet As ADODB.Recordset, VDBCon As ADODB.Connection, ByVal VScript)
Dim strSQL As String
If vRSet Is Nothing Then
   Set vRSet = New ADODB.Recordset
Else
   If vRSet.State = adStateOpen Then vRSet.Close
End If
strSQL = VScript
'    Debug.Print strSQL
vRSet.CursorLocation = adUseClient
'    Debug.Print strSQL
vRSet.Open strSQL, VDBCon, adOpenKeyset, adLockReadOnly
End Sub


Function SelisihHariJam(ByVal Awal As Date, _
                        ByVal Akhir As Date, _
                        RetValue As Integer) As String
                        
'parameter -> RetValue  1 = Hari
'                       2 = Jam
'                       3 = Menit
'                       4 = Detik
'                       5 = complete string

Dim Detik As Long, Hari As Long, Jam As Long
Dim JamLengkap As String
   
  If Awal > Akhir Then
     MsgBox "Tanggal dan waktu awal harus lebih kecil " & vbCrLf & _
            "dari pada tanggal dan waktu akhir", _
            vbCritical, "Peringatan"
     Exit Function
  End If
  
  'Tampung dalam durasi satuan terkecil, yaitu: DETIK
  Detik = DateDiff("s", Awal, Akhir)
  
  'Hitung jumlah jam dgn cara membagi 3600
  '(backslash ("\") supaya menghasilkan
  'nilai Integer tanpa pembulatan ke atas)
  Jam = Detik \ 3600
  
  'Jika jumlah jam lebih besar dari 23
  'artinya: lebih dari 1 hari
  If Jam > 23 Then
     
     'Hitung jumlah hari dgn car membagi 24
     '(backslash ("\") supaya menghasilkan
     'nilai integer tanpa pembulatan ke atas)
     Hari = Jam \ 24
     
     'Hitung Durasi Jam dalam hh:mm:ss
     JamLengkap = Format((Akhir - Awal), "hh:mm:ss")
  
  Else 'Jika jumlah jam <= 23
      
     Hari = 0   'maka jumlah hari = nol
      
     'Hitung Durasi Jam dalam hh:mm:ss
     JamLengkap = Format((Akhir - Awal), "hh:mm:ss")
  End If
  
  If Hari = 0 Then  'Jika jumlah hari = 0
     
     'Tampung hasil akhirnya
     SelisihHariJam = JamLengkap
  
  Else  'Jika jumlah hari > 0, tampilkan jumlah harinya
     
     'Tampung hasil akhirnya
     
  End If
  
  Select Case RetValue
      Case 1: SelisihHariJam = Str(Hari)
      Case 2:
         Hari = (Hari * 24) + Val(Format((Akhir - Awal), "hh"))
         SelisihHariJam = Str(Hari)
      Case 3:
         Hari = (Hari * 24 * 60) + (Val(Format((Akhir - Awal), "hh")) * 60) + Val(Format((Akhir - Awal), "mm"))
         SelisihHariJam = Str(Hari)
      Case 4:
         Hari = (Hari * 24 * 60 * 60) + (Val(Format((Akhir - Awal), "hh")) * 60 * 60) + (Val(Format((Akhir - Awal), "mm")) * 60) + Val(Format((Akhir - Awal), "ss"))
         SelisihHariJam = Str(Hari)
      Case 5: SelisihHariJam = Hari & " hari, " & JamLengkap
  End Select
  
  Exit Function

End Function

Public Function WeekNumber(InDate As Date) As Integer
  Dim DayNo As Integer
  Dim StartDays As Integer
  Dim StopDays As Integer
  Dim StartDay As Integer
  Dim StopDay As Integer
  Dim VNumber As Integer
  Dim ThurFlag As Boolean

  DayNo = InDate - DateSerial(Year(InDate), 1, 0) 'Days(InDate)
  StartDay = Weekday(DateSerial(Year(InDate), 1, 1)) - 1
  StopDay = Weekday(DateSerial(Year(InDate), 12, 31)) - 1
  ' Number of days belonging to first calendar week
  StartDays = 7 - (StartDay - 1)
  ' Number of days belonging to last calendar week
  StopDays = 7 - (StopDay - 1)
  ' Test to see if the year will have 53 weeks or not
  If StartDay = 4 Or StopDay = 4 Then ThurFlag = True Else ThurFlag = False
  VNumber = (DayNo - StartDays - 4) / 7
  ' If first week has 4 or more days, it will be calendar week 1
  ' If first week has less than 4 days, it will belong to last year's
  ' last calendar week
  If StartDays >= 4 Then
     WeekNumber = Fix(VNumber) + 2
  Else
     WeekNumber = Fix(VNumber) + 1
  End If
  ' Handle years whose last days will belong to coming year's first
  ' calendar week
  If WeekNumber > 52 And ThurFlag = False Then WeekNumber = 1
  ' Handle years whose first days will belong to the last year's
  ' last calendar week
  If WeekNumber = 0 Then
     WeekNumber = WeekNumber(DateSerial(Year(InDate) - 1, 12, 31))
  End If
End Function
