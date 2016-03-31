VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents RsMaster As DBQuick
Attribute RsMaster.VB_VarHelpID = -1
Private RsDetail As New DBQuick
Private LSqlMaster As String
Private LSqlDetail As String
Private LKey As String
Private LFieldName As String
Private LValue As String
Private LMTable As String
Private LDateField As String

'*** mendefinisikan nama table yang digunakan untuk update data approval ***'
Public Property Let MasterTable(ByVal TableName As String)
   LMTable = TableName
End Property

'*** Menghasilkan return value berupa strinf SQL Data Header ***'
Public Property Get StrSQLMaster() As String
       StrSQLMaster = LSqlMaster
End Property

'*** Menghasilkan return value berupa string SQL untuk data detail/ detail item dari Data Header ***'
Public Property Get StrSQLDetail() As String
       StrSQLDetail = LSqlDetail
End Property

'*** mendefinisikan string SQL untuk Master Data / data header ***'
Public Property Let StrSQLMaster(ByVal strSQL As String)
       LSqlMaster = strSQL
End Property

'*** Mendefinisikan string SQl untk Data detail yang akan dihubungkan dengan data header ***'
Public Property Let StrSQLDetail(ByVal strSQL As String)
       LSqlDetail = strSQL
End Property

'*** mendefinisikan key field yang digunakan untuk merelasikan anatara data header dg data detail ***'
Public Property Let Key(ByVal StrKey As String)
       LKey = StrKey
End Property

'*** procedure untuk menentukan nama field approval yang akan diberi nilai dan nilainya     ***'
'*** Parameter : FieldName -> adalah nama field Approval yang akan diupdate                 ***'
'***             Value     -> adalah nilai dari field approval yang akan dimasukkan         ***'
'***             DateField -> adalah parameter untuk menentukan nama filed tanggal approval ***'
Public Sub ApprovalField(FieldName As String, Value As String, DateField As String)
   LFieldName = FieldName
   LValue = Value
   LDateField = DateField
End Sub

Private Sub Class_Initialize()
   Set RsMaster = New DBQuick
End Sub

'*** binding data detail tiap kali ada perpindahan record data header ***'
Private Sub RsMaster_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   If RsMaster.DBRecordset.Recordcount > 0 Then
      Debug.Print LSqlDetail & " where " & LKey & "='" & RsMaster.DBRecordset.Fields(LKey)
      RsDetail.DBOpen LSqlDetail & " where " & LKey & "='" & RsMaster.DBRecordset.Fields(LKey) & "'", CNN
   End If
End Sub

'*** memberikan return value berupa recordset dari data detail ***'
Public Property Get DetailRecordset() As DBQuick
   Set DetailRecordset = RsDetail
End Property

'*** memberikan return value berupa detail recordset dari data header ***'
Public Property Get MasterRecordset() As DBQuick
   Set MasterRecordset = RsMaster
End Property

'*** procedure yang digunakan untuk eksekusi approval dengan meng update data yg sudah ditentukan ***'
Public Sub Approve()
   If MessageBox("Yakin Data Di Setujui ?", "Konfirmasi", msgYesNo, msgQuestion) = 1 Then
      SendDataToServer " update [" & LMTable & "] set [" & LFieldName & "]='" & LValue & "', [" & LDateField & "]='" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "' where [" & LKey & "]='" & RsMaster.DBRecordset.Fields(LKey) & "'"
   End If
End Sub

'*** procedure yang digunakan untuk eksekusi recordset header dan detail / untuk mengaktifkan class ini ***'
Public Sub Start()
   On Error Resume Next
   RsMaster.DBOpen LSqlMaster, CNN
   If RsMaster.Recordcount <> 0 Then RsDetail.DBOpen LSqlDetail & " where " & LKey & "='" & RsMaster.DBRecordset.Fields(LKey) & "'", CNN
   
End Sub

