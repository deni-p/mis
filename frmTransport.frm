VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmTransport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transportation"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTransport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   Tag             =   "Transporter Setting"
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5535
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   5550
      Left            =   0
      ScaleHeight     =   5550
      ScaleWidth      =   11205
      TabIndex        =   10
      Top             =   0
      Width           =   11205
      Begin VB.OptionButton OptType 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Shipping Transport"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Index           =   1
         Left            =   2820
         TabIndex        =   2
         Top             =   150
         Width           =   1965
      End
      Begin VB.OptionButton OptType 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Jasa Kurir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Index           =   0
         Left            =   225
         TabIndex        =   1
         Top             =   150
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "fax"
         Height          =   330
         Index           =   5
         Left            =   2220
         MaxLength       =   20
         TabIndex        =   8
         Tag             =   "Partner"
         Top             =   2190
         Width           =   3045
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Phone"
         Height          =   330
         Index           =   4
         Left            =   2220
         MaxLength       =   20
         TabIndex        =   7
         Tag             =   "Partner"
         Top             =   1845
         Width           =   3045
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Person"
         Height          =   330
         Index           =   3
         Left            =   2220
         MaxLength       =   50
         TabIndex        =   6
         Tag             =   "Partner"
         Top             =   1500
         Width           =   3045
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Address"
         Height          =   330
         Index           =   2
         Left            =   2220
         MaxLength       =   50
         TabIndex        =   5
         Tag             =   "Partner"
         Top             =   1155
         Width           =   5760
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Expedisi"
         Height          =   330
         Index           =   1
         Left            =   2220
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "Partner"
         Top             =   810
         Width           =   3765
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "ID"
         Height          =   330
         Index           =   0
         Left            =   2220
         MaxLength       =   16
         TabIndex        =   3
         Tag             =   "Partner"
         Top             =   465
         Width           =   1935
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2760
         Index           =   0
         Left            =   105
         TabIndex        =   9
         Tag             =   "Partner"
         Top             =   2625
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   4868
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16777215
         HeadLines       =   2
         RowHeight       =   16
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
            DataField       =   "ID"
            Caption         =   "ID"
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
            DataField       =   "Expedisi"
            Caption         =   "Expedisi"
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
            DataField       =   "Address"
            Caption         =   "Alamat"
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
            DataField       =   "Person"
            Caption         =   "Orang yg bisa dihubungi"
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
            DataField       =   "Phone"
            Caption         =   "Telp"
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
         BeginProperty Column05 
            DataField       =   "Fax"
            Caption         =   "Faximile"
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
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
         EndProperty
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   195
         X2              =   2310
         Y1              =   2505
         Y2              =   2505
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   180
         X2              =   2280
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   195
         X2              =   2280
         Y1              =   1815
         Y2              =   1815
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   195
         X2              =   2280
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   195
         X2              =   2355
         Y1              =   1125
         Y2              =   1125
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   180
         X2              =   2370
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Courier"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   5
         Left            =   225
         TabIndex        =   16
         Top             =   885
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Faksimili"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   4
         Left            =   225
         TabIndex        =   15
         Top             =   2265
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telp"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   14
         Top             =   1920
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Orang yg bisa di hubungi"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   13
         Top             =   1575
         Width           =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   12
         Top             =   1230
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   11
         Top             =   540
         Width           =   165
      End
   End
End
Attribute VB_Name = "frmTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyData As New clsMaster
Dim Mtype As String

Private Sub Form_Activate()
'If Me.WindowState = 0 Then If Me.WindowState = 0 Then Me.WindowState = 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
'Set Picture1.Picture = LoadResPicture(101, 0)
'HiasForm Picture1, Me

MyDDE.SetPermissions = aksess.MayDo("Transporter")

HiasFormManTell Picture2, Me
GridLayout
OptType(0).BackColor = Picture2.BackColor
OptType(1).BackColor = Picture2.BackColor
'OptType(0).ForeColor = Picture1.BackColor
'OptType(1).ForeColor = Picture1.BackColor
Mtype = "EXPEDISI"
Call OptType_Click(0)
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

'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
'GridLayout
'OptType(0).BackColor = Picture2.BackColor
'OptType(1).BackColor = Picture2.BackColor
'OptType(0).ForeColor = Picture1.BackColor
'OptType(1).ForeColor = Picture1.BackColor
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmTransport = Nothing
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
If MyDDE.IsSucces = False Then Exit Sub
Select Case AdReasonActiveDb
       Case tmbAddNew:
            PrepareString
            'mVarDataDc = True
            txtBox(0).Enabled = False
            If Mtype = "EXPEDISI" Then
               MyDDE.GetFieldByName("ID") = MyData.PrepareIndex(tmbExpedTransport, 5, "", "EX/")
            Else
               MyDDE.GetFieldByName("ID") = MyData.PrepareIndex(tmbShipTransport, 5, "", "SH/")
            End If
            txtBox(1).SetFocus
            OptType(0).Enabled = False
            OptType(1).Enabled = False
       Case tmbEdit:
            txtBox(0).Enabled = False
            'mVarDataDc = True
            txtBox(1).SetFocus
            OptType(0).Enabled = False
            OptType(1).Enabled = False
       Case tmbCancel, tmbSave, tmbDelete:
            OptType(0).Enabled = True
            OptType(1).Enabled = True
       Case tmbPrint:
            If Mtype = "EXPEDISI" Then
               CallRPTReport "Expedisi.Rpt"
            Else
               CallRPTReport "Shipment.Rpt"
            End If
       Case tmbDelete:
            
       Case Else: 'mVarDataDc = False
End Select
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
PrepareQuery
Err.Clear
End Sub


Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim mDel As New clsDelete
Select Case AdReasonActiveDb
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               If mDel.CekDelete(txtBox(0), reDelMasterTransport) = False Then
                  MyDDE.IsChildMemberReady = True
                  PrepareQuery
               Else
                  MyDDE.CancelTrans = True
                  MessageBox "Record (" & txtBox(0) & ") Sedang Dipakai Transaksi Lain." & vbCrLf & "Record Tidak Bisa DiHapus.", "Peringatan", msgOkOnly
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
End Sub

Private Sub OptType_Click(Index As Integer)
If Index = 0 Then
   Mtype = "EXPEDISI"
Else
   Mtype = "Ship"
End If
OpenDB
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
    Set .BindForm = frmTransport
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
    If Mtype = "EXPEDISI" Then
       .PrepareQuery = "Select * from [Transport] where type = 'Expedisi' Order By ID"
    Else
       .PrepareQuery = "Select * from [Transport] where type = 'Ship' Order By ID"
    End If
End With
End Sub

Private Sub PrepareQuery()
With MyDDE
    .PrepareAppend = " INSERT INTO [Transport] (ID, Expedisi, Address, Person, Phone, Fax, Type) " & _
                     " VALUES (N'" & ValidString(txtBox(0)) & "', N'" & ValidString(txtBox(1)) & "', N'" & ValidString(txtBox(2)) & "', N'" & ValidString(txtBox(3)) & "', N'" & ValidString(txtBox(4)) & "', N'" & ValidString(txtBox(5)) & "','" & Mtype & "')"
                     
    .PrepareUpdate = " UPDATE [Transport] Set [Expedisi] = N'" & ValidString(txtBox(1)) & "',[Address] = N'" & ValidString(txtBox(2)) & "',[Person] = N'" & ValidString(txtBox(3)) & "',[Phone] = N'" & ValidString(txtBox(4)) & "',[fax] = N'" & ValidString(txtBox(5)) & "' WHERE     (ID = N'" & ValidString(txtBox(0)) & "') and (type ='" & Mtype & "')"
                     
    .PrepareDelete = " DELETE FROM [Transport] WHERE   (ID = N'" & ValidString(txtBox(0)) & "') and (Type='" & Mtype & "') "
End With
End Sub

Private Sub PrepareString()
MyDDE.GetFieldByName(0) = "-"
MyDDE.GetFieldByName(1) = "-"
MyDDE.GetFieldByName(2) = "-"
MyDDE.GetFieldByName(3) = "-"
MyDDE.GetFieldByName(4) = "-"
MyDDE.GetFieldByName(5) = "-"
End Sub

Private Sub txtBox_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 4 Or Index = 5 Then
   ValidNum KeyAscii
End If
End Sub

Private Sub GridLayout()
DataGrid1(0).Height = 2760
DataGrid1(0).width = 10755
DataGrid1(0).Columns(0).width = 1604.976
DataGrid1(0).Columns(1).width = 2280.189
DataGrid1(0).Columns(2).width = 1980.284
DataGrid1(0).Columns(3).width = 1620.284
DataGrid1(0).Columns(4).width = 1349.858
DataGrid1(0).Columns(5).width = 1349.858
End Sub

