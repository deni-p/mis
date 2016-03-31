VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 2  'vbComplexBound
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function OpenBankAccount(ByVal GetPartnerBank As String) As Recordset
On Error Resume Next
CloseDB OpenBankAccount
Set OpenBankAccount = New Recordset
OpenBankAccount.CursorLocation = adUseClient
If GetPartnerBank = "" Then GetPartnerBank = "xxxxxxx"
   OpenBankAccount.Open "Select * from [Bank Partner] where PartnerID='" & GetPartnerBank & "'", CNN, adOpenStatic, adLockBatchOptimistic, adCmdText
End Function

Public Function OpenGudang() As Recordset
On Error Resume Next
CloseDB OpenGudang
Set OpenGudang = New Recordset
OpenGudang.CursorLocation = adUseClient
OpenGudang.Open "Select * from WareHouse order by WareHouse", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
Set OpenGudang.ActiveConnection = Nothing
End Function

Public Function OpenKelompok() As Recordset
On Error Resume Next
CloseDB OpenKelompok
Set OpenKelompok = New Recordset
OpenKelompok.CursorLocation = adUseClient
OpenKelompok.Open "Select * from [Inventory Group] order by NoGroup", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
Set OpenKelompok.ActiveConnection = Nothing
End Function

Public Function PrepareIndex(ByVal IndexTransaksi As TombolTransaksi, ByVal LenIdx As Integer, Optional ByVal ParamString As String, Optional ByVal IdxString As String)
Dim idx As New Recordset
Dim mVarNo As Long
Dim StrIdx As String
idx.CursorLocation = adUseClient
Select Case IndexTransaksi
       Case tmbInventory: StrIdx = " SELECT MAX(RIGHT(NoItem, 5)) AS MaxNom FROM Inventory WHERE     (StatusItem = N'ITEM')"
       Case tmbCustomer: StrIdx = " SELECT RIGHT(MAX(PartnerID), 7) AS MaxNom FROM PartnerDB GROUP BY PartnerType HAVING (PartnerType = N'CUSTOMER')"
       Case tmbSupplier: StrIdx = " SELECT RIGHT(MAX(PartnerID), 7) AS MaxNom FROM PartnerDB GROUP BY PartnerType HAVING (PartnerType = N'SUPPLIER')"
       Case tmbKelompok: StrIdx = " SELECT RIGHT(MAX(PartnerID), 7) AS MaxNom FROM PartnerDB GROUP BY PartnerType HAVING (PartnerType = N'SUPPLIER')"
       Case tmbShipTransport: StrIdx = "SELECT RIGHT(MAX(ID), 5) AS MaxNom FROM Transport GROUP BY Type HAVING (Type = N'SHIP')"
       Case tmbExpedTransport: StrIdx = "SELECT RIGHT(MAX(ID), 5) AS MaxNom FROM Transport GROUP BY Type HAVING (Type = N'EXPEDISI')"
End Select
idx.Open StrIdx, CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
Set idx.ActiveConnection = Nothing
With idx
     If .Recordcount <> 0 Then
        mVarNo = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + 1
     Else
        mVarNo = 1
     End If
     PrepareIndex = IdxString & "-" & KirimNull(LenIdx - Len(Trim(Str(mVarNo)))) & Trim(Str(mVarNo))
End With
CloseDB idx
End Function

