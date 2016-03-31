VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmBebanPembayaran 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Beban Pembayaran"
   ClientHeight    =   4125
   ClientLeft      =   1800
   ClientTop       =   2340
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBebanPembayaran.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   Tag             =   "Shipment"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
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
      Height          =   3570
      Left            =   0
      ScaleHeight     =   3570
      ScaleWidth      =   9255
      TabIndex        =   5
      Top             =   0
      Width           =   9255
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "TypeFreight"
         DataSource      =   "Adodc1"
         Height          =   330
         Index           =   0
         Left            =   1545
         MaxLength       =   25
         TabIndex        =   2
         Tag             =   "Partner"
         Top             =   495
         Width           =   3390
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "TypeLoco"
         DataSource      =   "Adodc1"
         Height          =   330
         Index           =   1
         Left            =   1545
         MaxLength       =   5
         TabIndex        =   1
         Tag             =   "Partner"
         Top             =   135
         Width           =   3390
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Beli"
         DataField       =   "Local"
         ForeColor       =   &H80000004&
         Height          =   300
         Left            =   5130
         TabIndex        =   3
         Tag             =   "Partner"
         Top             =   480
         Width           =   1515
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2505
         Index           =   0
         Left            =   105
         TabIndex        =   4
         Tag             =   "Partner"
         Top             =   915
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   4419
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   1
         BeginProperty Column00 
            DataField       =   "TypeFreight"
            Caption         =   "Local/Franco"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loco/Franco"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   540
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   6
         Top             =   210
         Width           =   165
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   120
         X2              =   3000
         Y1              =   435
         Y2              =   450
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   120
         X2              =   2400
         Y1              =   810
         Y2              =   810
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   3555
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmBebanPembayaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyID As New clsTransaksi

Private Sub Check1_Click()
   If Check1.Value = 0 Then
      Check1.Caption = "Beli"
   Else
      Check1.Caption = "Jual"
   End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()

MyDDE.SetPermissions = aksess.MayDo("Payment")

HiasFormManTell Picture2, Me
'HiasForm Picture1, Me

Check1.BackColor = Picture2.BackColor
GridLayout
OpenDB
DataGrid1(0).Columns(0).width = 7799.812
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MyDDE.CheckRecordPendinged = True Then
   ScanKey vbKeyF5, 0, MyDDE
   If MyDDE.IsSucces = True Then
      Cancel = False
      MyDDE.ClearRecordset
   Else
      Cancel = True
   End If
Else
   MyDDE.ClearRecordset
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next

Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmBebanPembayaran = Nothing
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 1
Select Case AdReasonActiveDb
       Case tmbAddNew:
            'mVarDataDc = True
            Kosong
            txtBox(0).SetFocus
       Case tmbEdit:
            txtBox(0).Enabled = False
            'mVarDataDc = True
            txtBox(1).Enabled = False
            txtBox(0).Enabled = True
            txtBox(0).SetFocus
       Case tmbPrint:
            'CallRPTReport "Tabel Mata Uang.rpt"
       Case Else: 'mVarDataDc = False
End Select
Exit Sub
1:
MessageBox Err.Description, "frmbebanpembayaran:mydde_afterprepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
PrepareQuery
Err.Clear
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo 2
If pRecordset.Recordcount <> 0 Then
   If Check1.Value = 0 Then
      Check1.Caption = "Beli"
   Else
      Check1.Caption = "Jual"
   End If
End If
Exit Sub
2:
MessageBox Err.Description, "frmbebanpembayaran:mydde_movecomplete" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 3
Dim mDel As New clsDelete
Select Case AdReasonActiveDb
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               If mDel.CekDelete(txtBox(0), reDelMasterBayar) = False Then
                  MyDDE.IsChildMemberReady = True
                  PrepareQuery
               Else
                  MyDDE.CancelTrans = True
                  MessageBox "Record (" & txtBox(0) & ") Sedang Dipakai Transaksi Lain." & vbCrLf & "Record Tidak Bisa DiHapus.", "Peringatan", msgOkOnly, msgCrtical
                  MyDDE.IsChildMemberReady = False
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               PrepareQuery
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
Set mDel = Nothing
Exit Sub
3:
MessageBox Err.Description, "frmbebanpembayaran:mydde_prepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub OpenDB()
With MyDDE
    .EditModeReplace = False
    Set .BindForm = frmBebanPembayaran
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT     * FROM         [Type Bayar]"
End With
End Sub

Private Sub PrepareQuery()
On Error GoTo xErr
With MyDDE
    .PrepareAppend = " INSERT INTO [Type Bayar] (TypeLoco,TypeFreight,[local]) " & _
                     " VALUES (N'" & txtBox(1) & "',N'" & ValidString(txtBox(0)) & "'," & Check1.Value & ")"
                     
    .PrepareUpdate = " UPDATE [Type Bayar] Set [local]=" & Check1.Value & ",[TypeFreight] = N'" & ValidString(txtBox(0)) & "' WHERE     (TypeLoco = N'" & ValidString(txtBox(1)) & "')"
                     
    .PrepareDelete = " DELETE FROM [Type Bayar] WHERE   (TypeLoco = N'" & ValidString(txtBox(1)) & "') "
End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub Kosong()
MyDDE.GetFieldByName("TypeLoco") = MyID.PrepareIndex(tmbFreight, 5, "", "")
MyDDE.GetFieldByName("TypeFreight") = "-"
txtBox(1).Enabled = False
End Sub

Private Sub GridLayout()
DataGrid1(0).Height = 2505
'DataGrid1(0).Width = 8355
End Sub
