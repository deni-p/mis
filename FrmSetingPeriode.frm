VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmSetingPeriode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setting Periode"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSetingPeriode.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   Tag             =   "Period Setting"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4830
      Left            =   0
      ScaleHeight     =   4830
      ScaleWidth      =   9345
      TabIndex        =   4
      Top             =   0
      Width           =   9345
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Kode Periode"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   0
         Left            =   1830
         TabIndex        =   0
         Tag             =   "SETING"
         Top             =   135
         Width           =   1935
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3570
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Tag             =   "SETING"
         Top             =   945
         Width           =   9090
         _ExtentX        =   16034
         _ExtentY        =   6297
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
            Weight          =   700
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Kode Periode"
            Caption         =   "No.Periode"
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
            DataField       =   "Tanggal Mulai"
            Caption         =   "Tanggal Mulai"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd MMMM yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Tutup Buku"
            Caption         =   "Status"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Closed"
               FalseValue      =   "Open"
               NullValue       =   "Open"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Periode"
            Caption         =   "Periode"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "grace_period"
            Caption         =   "Grace Period (hari)"
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
               Alignment       =   2
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "Tanggal Mulai"
         Height          =   315
         Left            =   1815
         TabIndex        =   1
         Tag             =   "SETING"
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
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
      Begin VB.Line Line1 
         Index           =   1
         X1              =   225
         X2              =   2055
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   210
         X2              =   2040
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Awal"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   3
         Top             =   540
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No.Periode"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   5
         Top             =   195
         Width           =   795
      End
   End
End
Attribute VB_Name = "FrmSetingPeriode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mVarThn, mVarBln As String
Dim mVarTahun As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
MyDDE.SetPermissions = aksess.MayDo("Periode Transaksi") 'Set Akses Tombol
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmSetingPeriode
    .BindFormTAG = "SETING"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT GlFile AS [Kode Periode], BeginDate AS [Tanggal Mulai], EndDate AS [Tanggal Selesai], " & _
    " Closed AS [Tutup Buku], Periode, grace_period FROM  SettingPeriod ORDER BY GlFile"
End With
mVarTahun = MakeTahun(True)
Dim Mydas As Integer
'messagebox CDate(DateSerial(DTPicker1.Year, DTPicker1.Month + 1, DTPicker1.Day))
'Mydas = CDate(DateSerial(DTPicker1.Year, DTPicker1.Month + 1, DTPicker1.Day)) - DTPicker1.Value
'messagebox Mydas - 1
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
GridLayout
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmSetingPeriode = Nothing
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim mPer As Byte
'Dim mDaysOfMonth As Integer
Select Case AdReasonActiveDb
       Case tmbAddNew:
            txtBox(0).Enabled = False
            DTPicker1.SetFocus
            mPer = MakePeriode
            mVarThn = Trim(Str(mVarTahun)) & "-" & Format(mPer, "0#")
            'DTPicker1.Value = DateSerial(mVarTahun, mPer, 1)
'            MyDDE.GetFieldByName("Tanggal Mulai") = DateSerial(mVarTahun, mPer, 1)
'            mDaysOfMonth = (CDate(DateSerial(DTPicker1.Year, DTPicker1.Month + 1, DTPicker1.Day)) - DTPicker1.Value) - 1
            'MyDDE.GetFieldByName("Tanggal Selesai") = DateSerial(DTPicker1.Year, DTPicker1.Month, DTPicker1.Day + mDaysOfMonth)
            'messagebox MyDDE.GetFieldByName("Tanggal Selesai")
            DTPicker1.Value = CDate(Format(dDateBegin, "dd/mm/yyyy"))
            MyDDE.GetFieldByName("Kode Periode") = mVarThn
            MyDDE.GetFieldByName("Tutup Buku") = False
            MyDDE.GetFieldByName("Periode") = Val(Right(mVarThn, 2)) 'Trim(Mid(GlFileIndex, 6, 5))
            MyDDE.GetFieldByName("Tanggal Mulai") = DTPicker1.Value
       Case tmbEdit:
            txtBox(0).Enabled = False
            'mVarDataDc = True
            DTPicker1.SetFocus
       Case tmbPrint:
            CallRPTReport "Seting Periode.rpt"
       Case Else: 'mVarDataDc = False
End Select
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim mDel As New clsDelete
Select Case AdReasonActiveDb
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               If mDel.CekDelete(txtBox(0), reDelMasterCurency) = False Then
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

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub


Private Sub PrepareQuery()
Dim mDaysOfMonth As Integer
mDaysOfMonth = (CDate(DateSerial(DTPicker1.Year, DTPicker1.Month + 1, DTPicker1.Day)) - DTPicker1.Value) - 1
'MyDDE.GetFieldByName("Tanggal Selesai") = DateSerial(DTPicker1.Year, DTPicker1.Month, DTPicker1.Day + mDaysOfMonth)
With MyDDE
    .PrepareAppend = " INSERT INTO SettingPeriod" & _
                     " (GlFile, BeginDate, EndDate, Closed, Periode)" & _
                     " VALUES     (N'" & txtBox(0) & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format(DateSerial(DTPicker1.Year, DTPicker1.Month, DTPicker1.Day + mDaysOfMonth), "dd/mm/yy") & "', 3), 0, " & Val(Right(mVarThn, 2)) & ")"

    .PrepareUpdate = " UPDATE    SettingPeriod" & _
                     " SET  BeginDate = CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), EndDate = CONVERT(DATETIME, '" & Format(DateSerial(DTPicker1.Year, DTPicker1.Month, DTPicker1.Day + mDaysOfMonth), "dd/mm/yy") & "', 3)" & _
                     " WHERE     (GlFile = N'" & txtBox(0) & "')"

    .PrepareDelete = " DELETE FROM SettingPeriod WHERE     (GlFile = N'" & txtBox(0) & "')"
End With
End Sub

Private Function MakePeriode() As Byte
Dim Rc As New DBQuick
Rc.DBOpen "SELECT     MAX(Periode) AS MaxOfPeriode FROM         SettingPeriod WHERE     (LEFT(GlFile, 4) = '" & mVarTahun & "')", CNN, lckLockReadOnly
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        MakePeriode = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        MakePeriode = 0
     End If
     MakePeriode = MakePeriode + 1
     If MakePeriode = 13 Then
        mVarTahun = MakeTahun
        MakePeriode = 1
     End If
End With
Rc.CloseDB
End Function

Private Function MakeTahun(Optional ByVal Tipical As Boolean) As Long
Dim Rc As New DBQuick
Rc.DBOpen "SELECT     MAX(LEFT(GlFile, 4)) AS MaxOfGlFile FROM         SettingPeriod", CNN, lckLockReadOnly
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        If Tipical = False Then
           MakeTahun = Val(IIf(Not IsNull(.Fields(0)), .Fields(0), mVarTahun)) + 1
        Else
           MakeTahun = Val(IIf(Not IsNull(.Fields(0)), .Fields(0), mVarTahun))
        End If
     Else
        MakeTahun = Val(Right(dDateBegin, 4))
     End If
End With
Rc.CloseDB
End Function

Private Sub GridLayout()
DataGrid1(0).Columns(0).width = 1995.024
DataGrid1(0).Columns(1).width = 2300
DataGrid1(0).Columns(2).width = 1500
DataGrid1(0).Columns(3).width = 1000
DataGrid1(0).Columns(4).width = 1950

End Sub


