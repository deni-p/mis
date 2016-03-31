VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmMemorial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Jurnal Umum"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10905
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMemorial.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   Tag             =   "Memorial Journal"
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5340
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   0
      ScaleHeight     =   5415
      ScaleWidth      =   10905
      TabIndex        =   6
      Top             =   0
      Width           =   10905
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         DataField       =   "Keterangan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1455
         MaxLength       =   200
         TabIndex        =   3
         Tag             =   "ASM"
         Top             =   840
         Width           =   6615
      End
      Begin MSDataGridLib.DataGrid DGPurchase 
         Bindings        =   "frmMemorial.frx":6852
         Height          =   3480
         Left            =   105
         TabIndex        =   4
         Top             =   1335
         Width           =   10620
         _ExtentX        =   18733
         _ExtentY        =   6138
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16777215
         HeadLines       =   2
         RowHeight       =   15
         RowDividerStyle =   6
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "No Akun"
            Caption         =   "No Akun"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Nama Akun"
            Caption         =   "Nama Akun"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Keterangan"
            Caption         =   "Keterangan"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Doc Reff"
            Caption         =   "Doc Reff"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Debet"
            Caption         =   "Debet"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0;(#,##0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Kredit"
            Caption         =   "Kredit"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0;(#,##0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "Tanggal"
         Height          =   315
         Left            =   1455
         TabIndex        =   2
         Tag             =   "ASM"
         Top             =   450
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   556
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
         CurrentDate     =   38272
      End
      Begin VB.Label LblAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8250
         TabIndex        =   5
         Top             =   4875
         Width           =   2460
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   3
         Left            =   6960
         TabIndex        =   10
         Top             =   4935
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Bukti"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   9
         Top             =   128
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   8
         Top             =   510
         Width           =   570
      End
      Begin VB.Label lblFixAssets 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No Bukti"
         DataField       =   "No Bukti"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   1455
         TabIndex        =   1
         Tag             =   "ASM"
         Top             =   60
         Width           =   2220
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   7
         Top             =   915
         Width           =   840
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   105
         X2              =   1620
         Y1              =   1155
         Y2              =   1155
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   105
         X2              =   1620
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   6915
         X2              =   8430
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   90
         X2              =   1605
         Y1              =   375
         Y2              =   375
      End
   End
End
Attribute VB_Name = "frmMemorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RcGroup As New DBQuick
Private MyData As New clsTransaksi
Private mVarAdd, mFirstCaller As Boolean
Private mVarTmp As String
Private RcPartner As New DBQuick
Dim IDGen As New IDGenerator

Private Sub CmdDocReff_Click()
OpenPartner 1
End Sub

Private Sub DGPurchase_AfterColEdit(ByVal ColIndex As Integer)
Select Case DGPurchase.col
       Case 4:
             If DGPurchase.Columns(4) = "" Then DGPurchase.Columns(4) = "0"
             If DGPurchase.Columns(4).Value <> 0 Then DGPurchase.Columns(5).Value = 0
             Totaldata
       Case 5:
             If DGPurchase.Columns(5) = "" Then DGPurchase.Columns(5) = "0"
             If DGPurchase.Columns(5).Value <> 0 Then DGPurchase.Columns(4).Value = 0
             Totaldata
       Case Else: If DGPurchase.Columns(ColIndex) = "" Then DGPurchase.Columns(ColIndex) = "-"
End Select
End Sub

Private Sub DGPurchase_Error(ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub

Private Sub DGPurchase_KeyDown(KeyCode As Integer, Shift As Integer)
If mVarAdd = False Then Exit Sub
Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If mVarAdd = True Then
   DGPurchase.MarqueeStyle = dbgFloatingEditor
   Select Case DGPurchase.col
          Case 2, 3, 4, 5: DGPurchase.AllowUpdate = True
'               If DGPurchase.Col = 3 Then
'                  MoveButton True
'               Else
'                  MoveButton False
'               End If
          Case Else: DGPurchase.AllowUpdate = False
'          MoveButton False
   End Select
Else
   DGPurchase.AllowUpdate = False
   DGPurchase.MarqueeStyle = dbgHighlightRow
End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
GridLayout
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
Set mCall = New frmCaller
DTPicker1.Value = Date
With MyDDE
    .EditModeReplace = False
    Set .BindForm = frmMemorial
    .BindFormTAG = "ASM"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT JournalID as [No Bukti], DateTrans as [Tanggal], Note as [Keterangan] FROM         [Table Journal] WHERE     (TypeTrans = N'MEMORIAL') AND (Status = 0)"
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
RcGroup.CloseDB
MyDDE.ClearRecordset
Set MyData = Nothing
Set mCall = Nothing
End Sub

Private Sub Form_Resize()

'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmMemorial = Nothing
End Sub

Private Sub mCall_BeforeUnload()
If mCall.FromTagActive = "MASTER PERKIRAAN" Then
   If DGPurchase.Enabled = True Then DGPurchase.SetFocus
   mFirstCaller = False
End If
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case TagForm
       Case "MASTER PERKIRAAN":
            With MyDDE.ChildRecordset
                 .Fields("No Akun") = mCall.GetFieldByName(0)
                 .Fields("Nama Akun") = mCall.GetFieldByName(1)
                 .Fields("Keterangan") = "-"
                 .Fields("Doc Reff") = "-"
                 .Fields("Debet") = 0
                 .Fields("Kredit") = 0
            End With
       Case "MASTER KAS":
            With MyDDE.ActiveRecordset
                 .Fields("BankID") = mCall.GetFieldByName(0)
                 .Fields("NamaBank") = mCall.GetFieldByName(1)
            End With
            'TotalKas NoVoucher(1)
       Case "MASTER AKTIVA TETAP":
            With MyDDE.ChildRecordset
                 .Fields(0) = mCall.GetFieldByName(0)
                 .Fields(1) = mCall.GetFieldByName(1)
                 .Fields(2) = 0
                 .Fields(3) = 0
            End With
       Case Else
            Select Case mVarTmp
                   Case "KODE SUPPLIER": MyDDE.ChildRecordset.Fields(3) = mCall.GetFieldByName(0)
                   Case "KODE CUSTOMER": MyDDE.ChildRecordset.Fields(3) = mCall.GetFieldByName(0)
                   Case "KODE KARYAWAN": MyDDE.ChildRecordset.Fields(3) = mCall.GetFieldByName(0)
                   Case "KODE BARANG": MyDDE.ChildRecordset.Fields(3) = mCall.GetFieldByName(0)
                   Case "KODE AKTIVA": MyDDE.ChildRecordset.Fields(3) = mCall.GetFieldByName(0)
                   Case "KODE KAS": MyDDE.ChildRecordset.Fields(3) = mCall.GetFieldByName(0)
            End Select
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit:
            'txtBox(0).Enabled = False
            mVarAdd = True
            DTPicker1.SetFocus
       Case tmbAddNew:
            DTPicker1.Value = CDate(Format(Date, "dd/mm/yyyy"))
            With MyDDE
'                 .GetFieldByName("No Bukti") = MyData.PrepareIndex(tmbTransaksiMemorial, 5, "", TglIndex)
                 .GetFieldByName("No Bukti") = IDGen.GetID("MM")
                 .GetFieldByName("Keterangan") = "Transaksi Memorial"
                 .GetFieldByName("Tanggal") = DTPicker1.Value
            End With
            mVarAdd = True
            DTPicker1.SetFocus
       Case tmbDetail:
            mVarAdd = True
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               If OpenPartner(0) = True Then CancelDetailTrans
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbDelete:
            If MyDDE.IsChildMemberReady = True Then
               'SendDataToServer ("DELETE FROM Inventory WHERE     (NoItem = N'" & 'txtBox(0) & "') ")
            End If
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               SimpanDetail
            End If
       Case tmbPrint:
            CallRPTReport "Bukti Memorial.Rpt", "Select * from [bukti memorial] Where [No Bukti]='" & lblFixAssets(0) & "'"
            
'       Case Else: 'mVarDataDc = False
End Select
'cmdLink(0).Enabled = mVarAdd
'cmdLink(1).Enabled = mVarAdd
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("No Bukti")), MyDDE.GetFieldByName("No Bukti"), "XXXXX")
Totaldata 'TotalKas IIf(Not IsNull(MyDDE.GetFieldByName("BankID")), MyDDE.GetFieldByName("BankID"), "XXXXX")
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbSave:
'            If MyDDE.IsChildMemberReady = True Then
                If (MyDDE.ChildRecordset.Recordcount <> 0) Then
                   If (MyDDE.ChildRecordset.Fields(4) <> 0 Or MyDDE.ChildRecordset.Fields(5) <> 0) Then
                      MyDDE.CancelTrans = TotalKas
                      If MyDDE.CancelTrans = True Then
                         MyDDE.IsChildMemberReady = False
                         MessageBox "Data detail belum Balance. Harap diperiksa dulu.", "Peringatan", msgOkOnly, msgCrtical
                      Else
                         MyDDE.IsChildMemberReady = True
'                         MessageBox "Data detail belum Balance. Harap diperiksa dulu.", "Peringatan", msgOkOnly
                         'SimpanDetail
                         mVarAdd = False
                      End If
                   Else
                      MyDDE.IsChildMemberReady = False
                      MessageBox "Data detail belum Balance. Harap diperiksa dulu.", "Peringatan", msgOkOnly, msgCrtical
                   End If
                Else
                   MessageBox "Data detail belum Ada. Harap diperiksa dulu.", "Peringatan", msgOkOnly, msgCrtical
                   MyDDE.CancelTrans = True
                End If
'            Else
'               MessageBox "Data detail belum Lengkap. Harap diisi dulu.", "Peringatan", msgOkOnly
'            End If
'            If MyDDE.CheckEmptyControl = False Then
'               If MyDDE.ChildRecordset.Recordcount <> 0 Then
'                  MyDDE.IsChildMemberReady = True
'               Else
'                  MyDDE.IsChildMemberReady = False
'                  MessageBox "Data detail belum ada. Harap diisi dulu.", "Peringatan", msgOkOnly
'               End If
'            Else
'               MyDDE.IsChildMemberReady = False
'            End If
       Case tmbDetail:
            MyDDE.CancelTrans = mFirstCaller
            If MyDDE.CancelTrans = True Then Exit Sub
               If MyDDE.ChildRecordset.Recordcount <> 0 Then
                  If MyDDE.ChildRecordset.Fields(4) = 0 And MyDDE.ChildRecordset.Fields(5) = 0 Then
                     MyDDE.IsChildMemberReady = False
                     MyDDE.CancelTrans = True
                     MessageBox "Jumlah transaksi harus isi.", "Peringatan", msgOkOnly, msgCrtical
                  Else
                     MyDDE.IsChildMemberReady = True
                     MyDDE.CancelTrans = False
                  End If
               Else
                  MyDDE.IsChildMemberReady = True
                  MyDDE.CancelTrans = False
               End If
       
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
            Else
               MyDDE.IsChildMemberReady = False
            End If
            mVarAdd = False
       Case tmbCancel:
            mVarAdd = False
'       Case tmbDetail:
'            If MyDDE.CheckEmptyControl = False Then
'               MyDDE.IsChildMemberReady = True
'               OpenPartner
'            Else
'               MyDDE.IsChildMemberReady = False
'            End If
       Case tmbSave:
               If MyDDE.ChildRecordset.Recordcount <> 0 Then
                  If (MyDDE.ChildRecordset.Fields(4) = 0 Or MyDDE.ChildRecordset.Fields(5) = 0) And MyDDE.ChildRecordset.Fields(5) = 0 Then
                     MyDDE.IsChildMemberReady = False
                     MyDDE.CancelTrans = True
                     MessageBox "Jumlah transaksi harus isi.", "Peringatan", msgOkOnly, msgCrtical
                  Else
                     MyDDE.IsChildMemberReady = True
                     MyDDE.CancelTrans = False
                  End If
               Else
                  MyDDE.IsChildMemberReady = True
                  MyDDE.CancelTrans = False
               End If
End Select
End Sub

Private Sub PrepareQuery()
On Error Resume Next
With MyDDE
    .PrepareAppend = " INSERT INTO [Table Journal]" & _
                     " (JournalID, DateTrans, Periode, TypeTrans, Note,NoUrut) " & _
                     " VALUES     (N'" & lblFixAssets(0) & "','" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'," & mVarPeriode & ", N'MEMORIAL', N'" & ValidString(txtNote.Text) & "','" & MyData.PrepareIndex(tmbTransaksiNOJOURNAL, 13, Format(Year(dDateBegin), "yyyy"), "JR" & Format(Year(dDateBegin), "yyyy") & "-") & "')"

    .PrepareUpdate = " UPDATE [Table Journal]" & _
                     " SET DateTrans = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "', Periode = " & mVarPeriode & ", TypeTrans = N'MEMORIAL'," & _
                     " Note = N'" & ValidString(txtNote) & "'" & _
                     " WHERE     (JournalID = N'" & lblFixAssets(0) & "')"

    .PrepareDelete = " DELETE FROM [Table Journal] WHERE     (JournalID = N'" & lblFixAssets(0) & "') "
End With
Err.Clear
End Sub

Private Sub SimpanDetail()
With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
        If SendDataToServer("DELETE FROM [Detail Journal] WHERE     ([JournalID] = N'" & lblFixAssets(0) & "')") = True Then
           .MoveFirst
           Do
             If .EOF = True Then Exit Do
             SendDataToServer (" INSERT INTO [Detail Journal] " & _
                               " (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan)" & _
                               " VALUES (N'" & lblFixAssets(0) & "', N'" & .Fields("No Akun") & "', N'" & .Fields("Doc Reff") & "', " & CCur(.Fields("Debet")) & ", " & .Fields("Kredit") & ", N'" & Left(.Fields("Keterangan"), 244) & "')")
                               
             'SendDataToServer (" UPDATE    [Tabel Pembantu]" & _
                               " SET CurrentDR" & mVarPeriode & " =  CurrentDR" & mVarPeriode & " + " & CCur(.Fields("Debet")) & ", CurrentCR" & mVarPeriode & " = CurrentCR" & mVarPeriode & " + " & CCur(.Fields("kredit")) & _
                               " WHERE     (NoAccount = N'" & .Fields("No Akun") & "')")
                               
             .MoveNext
           Loop
           .MoveLast
        End If
     End If
End With
End Sub

Private Sub OpenDetail(ByVal ParamString As String)
Dim RcDetail As New DBQuick
RcDetail.DBOpen " SELECT     [Detail Journal].NoAccount AS [No Akun], GLAccount.AccountName AS [Nama Akun], [Detail Journal].Keterangan, [Detail Journal].[Doc Reff],                        [Detail Journal].Debet AS Debet, [Detail Journal].Credit AS Kredit FROM         [Detail Journal] INNER JOIN                       GLAccount ON [Detail Journal].NoAccount = GLAccount.NoAccount WHERE     ([Detail Journal].JournalID = N'" & ParamString & "') ORDER BY [Detail Journal].Debet DESC ", CNN
Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
Set DGPurchase.DataSource = MyDDE.ChildRecordset
RcDetail.CloseDB
End Sub

Private Function OpenPartner(ByVal Index As Integer) As Boolean

mVarTmp = ""
Select Case Index
       Case 0:
            RcPartner.DBOpen "SELECT     NoAccount AS [No Akun], AccountName AS [Nama Akun] FROM         GLAccount WHERE     ([Group] = N'Detail List Account') ORDER BY NoAccount", CNN, lckLockReadOnly
            mFirstCaller = True
       Case 1:
            mVarTmp = UCase(OpenDocReff)
            Select Case mVarTmp
                   Case "KODE SUPPLIER":
                         RcPartner.DBOpen " SELECT     PartnerID AS [Kode Supplier], CompanyName AS [Nama Perusahaan] FROM         PartnerDB WHERE     (PartnerType = N'SUPPLIER') ORDER BY PartnerID", CNN, lckLockReadOnly
                   Case "KODE CUSTOMER":
                         RcPartner.DBOpen " SELECT     PartnerID AS [Kode Customer], CompanyName AS [Nama Perusahaan] FROM         PartnerDB WHERE     (PartnerType = N'CUSTOMER') ORDER BY PartnerID", CNN, lckLockReadOnly
                   Case "KODE KARYAWAN":
                         RcPartner.DBOpen " SELECT     EmpID AS [Kode Karyawan], FullName AS [Nama Karyawan] FROM         Employees ORDER BY EmpID", CNN, lckLockReadOnly
                   Case "KODE BARANG":
                         RcPartner.DBOpen " SELECT     NoItem AS [Kode Barang], ItemName AS [Nama Barang] FROM         Inventory ORDER BY NoItem", CNN, lckLockReadOnly
                   Case "KODE AKTIVA":
                         RcPartner.DBOpen " SELECT     NoItem AS [Kode Barang], ItemName AS [Nama Barang] FROM         Inventory ORDER BY NoItem", CNN, lckLockReadOnly
                   Case "KODE KAS"
                         RcPartner.DBOpen "SELECT     BankID, NamaBank FROM         [Temp Bank] ORDER BY BankID", CNN, lckLockReadOnly
            End Select

       Case 2:
            RcPartner.DBOpen "SELECT     [No Aktiva], [Nama Aktiva] FROM         [Tabel Aktiva Tetap] ORDER BY [No Aktiva]", CNN, lckLockReadOnly
            
End Select
If RcPartner.Recordcount <> 0 Then
    Select Case Index
           Case 0: mCall.FromTagActive = "MASTER PERKIRAAN"
           Case 1: mCall.FromTagActive = mVarTmp
           Case 2: mCall.FromTagActive = "MASTER AKTIVA TETAP"
    End Select
    Set mCall.FormData = RcPartner.DBRecordset
    mCall.LookUp Me
'    If FindOwnRecordset(MyDDE.ChildRecordset, "[No Akun] = '" & MyDDE.ChildRecordset.Fields(0) & "'") = True Then
'       MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("No Akun") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
'        MyDDE.ChildRecordset.CancelBatch adAffectCurrent
'        If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
'       DGPurchase.SetFocus
'    End If
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgCrtical
   OpenPartner = True
End If
End Function

Private Sub CancelDetailTrans()
If MyDDE.ChildRecordset.Recordcount <> 0 Then
  If Not MyDDE.ChildRecordset.EOF Then MyDDE.ChildRecordset.MoveNext
  If MyDDE.ChildRecordset.EOF And MyDDE.ChildRecordset.Recordcount > 0 Then MyDDE.ChildRecordset.MoveLast
End If
End Sub

Private Function TglIndex() As String
Dim TglHari As String
Dim TglBulan As String
Dim TglTahun As String
TglIndex = "MM-" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End Function


Private Function TotalKas() As Boolean
Dim RcKas As New DBQuick
Dim mVarData As Variant
Dim mDr As Variant
Dim mCr As Variant
Dim I As Long
Set RcKas.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
mDr = 0
mCr = 0
With RcKas.DBRecordset
     If .Recordcount <> 0 Then
        mVarData = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(mVarData, 2)
            mDr = mDr + IIf(Not IsNull(mVarData(4, I)), mVarData(4, I), 0)
            mCr = mCr + IIf(Not IsNull(mVarData(5, I)), mVarData(5, I), 0)
        Next I
        If (mDr - mCr) <> 0 Then
           TotalKas = True
        End If
     End If
End With
RcKas.CloseDB
End Function

Private Sub Totaldata()
Dim RcKas As New DBQuick
Dim mVarData As Variant
Dim I As Long
Set RcKas.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
LblAmount = 0
With RcKas.DBRecordset
     If .Recordcount <> 0 Then
        mVarData = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(mVarData, 2)
            LblAmount = FormatNumber(LblAmount + IIf(Not IsNull(mVarData(4, I)), mVarData(4, I), 0), 0)
        Next I
     End If
End With
RcKas.CloseDB
End Sub

'Private Sub MoveButton(Optional ByVal Tipical As Boolean)
'CmdDocReff.Enabled = Tipical
'CmdDocReff.Visible = Tipical
'If Tipical = True Then
'   With DGPurchase
'        CmdDocReff.Move (.Columns(3).Left + .Columns(3).Width) - 200, (.RowTop(.Row) + .RowHeight) + 715, 300, .RowHeight
'   End With
'End If
'End Sub

Private Function OpenDocReff() As String
Dim AA As New DBQuick
If MyDDE.ChildRecordset.Recordcount <> 0 Then
    AA.DBOpen " SELECT     NoAccount, [Value Data] FROM         [Daftar Configurasi] GROUP BY NoAccount, [Value Data] HAVING      (NoAccount = N'" & MyDDE.ChildRecordset.Fields("No Akun") & "')", CNN, lckLockReadOnly
    With AA.DBRecordset
         If .Recordcount <> 0 Then
             OpenDocReff = IIf(Not IsNull(.Fields(1)), .Fields(1), "")
         End If
    End With
    AA.CloseDB
'Else
'   CmdDocReff.Enabled = False
'   CmdDocReff.Visible = False
End If
End Function

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

Private Sub GridLayout()
DGPurchase.Columns(0).width = 1395.213
DGPurchase.Columns(1).width = 2174.74
DGPurchase.Columns(2).width = 2220.094
DGPurchase.Columns(3).width = 1244.976
DGPurchase.Columns(4).width = 1514.835
DGPurchase.Columns(5).width = 1514.835
End Sub


Private Sub txtNote_GotFocus()
Block txtNote
End Sub
