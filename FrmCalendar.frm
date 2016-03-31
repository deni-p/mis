VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmCalendar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calendar"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10065
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCalendar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   Tag             =   "Scheduling Calendar"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   5310
      Left            =   0
      ScaleHeight     =   5310
      ScaleWidth      =   10065
      TabIndex        =   13
      Top             =   0
      Width           =   10065
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Calendar ID"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   0
         Left            =   1335
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Partner"
         Top             =   255
         Width           =   1965
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Keterangan"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   1
         Left            =   1335
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "Partner"
         Top             =   615
         Width           =   3045
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAAF6F&
         Caption         =   " Configuration "
         Height          =   1215
         Left            =   150
         TabIndex        =   14
         Top             =   1005
         Width           =   9780
         Begin VB.CheckBox DaysCek 
            Appearance      =   0  'Flat
            BackColor       =   &H00EAAF6F&
            Caption         =   "Sun"
            DataField       =   "Day1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000004&
            Height          =   240
            Index           =   0
            Left            =   780
            TabIndex        =   3
            Tag             =   "Partner"
            Top             =   330
            Width           =   780
         End
         Begin VB.CheckBox DaysCek 
            Appearance      =   0  'Flat
            BackColor       =   &H00EAAF6F&
            Caption         =   "Mon"
            DataField       =   "Day2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000004&
            Height          =   240
            Index           =   1
            Left            =   1665
            TabIndex        =   4
            Tag             =   "Partner"
            Top             =   330
            Width           =   780
         End
         Begin VB.CheckBox DaysCek 
            Appearance      =   0  'Flat
            BackColor       =   &H00EAAF6F&
            Caption         =   "Tue"
            DataField       =   "Day3"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000004&
            Height          =   240
            Index           =   2
            Left            =   2565
            TabIndex        =   5
            Tag             =   "Partner"
            Top             =   330
            Width           =   780
         End
         Begin VB.CheckBox DaysCek 
            Appearance      =   0  'Flat
            BackColor       =   &H00EAAF6F&
            Caption         =   "Wed"
            DataField       =   "Day4"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000004&
            Height          =   240
            Index           =   3
            Left            =   3450
            TabIndex        =   6
            Tag             =   "Partner"
            Top             =   330
            Width           =   780
         End
         Begin VB.CheckBox DaysCek 
            Appearance      =   0  'Flat
            BackColor       =   &H00EAAF6F&
            Caption         =   "Thu"
            DataField       =   "Day5"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000004&
            Height          =   240
            Index           =   4
            Left            =   4335
            TabIndex        =   7
            Tag             =   "Partner"
            Top             =   330
            Width           =   780
         End
         Begin VB.CheckBox DaysCek 
            Appearance      =   0  'Flat
            BackColor       =   &H00EAAF6F&
            Caption         =   "Fri"
            DataField       =   "Day6"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000004&
            Height          =   240
            Index           =   5
            Left            =   5235
            TabIndex        =   8
            Tag             =   "Partner"
            Top             =   330
            Width           =   780
         End
         Begin VB.CheckBox DaysCek 
            Appearance      =   0  'Flat
            BackColor       =   &H00EAAF6F&
            Caption         =   "Sat"
            DataField       =   "Day7"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000004&
            Height          =   240
            Index           =   6
            Left            =   6120
            TabIndex        =   9
            Tag             =   "Partner"
            Top             =   330
            Width           =   780
         End
         Begin MSComCtl2.DTPicker Tanggal 
            DataField       =   "Dari Tanggal"
            Height          =   300
            Index           =   0
            Left            =   1915
            TabIndex        =   10
            Tag             =   "Partner"
            Top             =   757
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   529
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
            CustomFormat    =   "dd MMM yyyy"
            Format          =   71630851
            CurrentDate     =   38461
         End
         Begin MSComCtl2.DTPicker Tanggal 
            DataField       =   "Sampai Tanggal"
            Height          =   300
            Index           =   1
            Left            =   4980
            TabIndex        =   11
            Tag             =   "Partner"
            Top             =   757
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   529
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
            CustomFormat    =   "dd MMM yyyy"
            Format          =   71630851
            CurrentDate     =   38461
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date"
            Height          =   195
            Index           =   0
            Left            =   780
            TabIndex        =   16
            Top             =   810
            Width           =   750
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End Date"
            Height          =   195
            Index           =   1
            Left            =   3935
            TabIndex        =   15
            Top             =   810
            Width           =   660
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2625
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   2340
         Width           =   9810
         _ExtentX        =   17304
         _ExtentY        =   4630
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16577005
         ForeColor       =   7159830
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "DateFrom"
            Caption         =   "Dari Tanggal"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd MMM yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "DateTo"
            Caption         =   "Sampai Tanggal"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd MMM yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Description"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               DividerStyle    =   6
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   6
            EndProperty
            BeginProperty Column02 
            EndProperty
         EndProperty
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   1
         X1              =   4365
         X2              =   165
         Y1              =   915
         Y2              =   915
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   3270
         X2              =   165
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Calendar ID"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   18
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   17
         Top             =   675
         Width           =   795
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5295
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mAdd As Boolean

Private Sub DataGrid1_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo 1
If mAdd = True Then
   DataGrid1(0).MarqueeStyle = dbgFloatingEditor
   DataGrid1(0).AllowUpdate = mAdd
Else
   DataGrid1(0).MarqueeStyle = dbgHighlightRow
   DataGrid1(0).AllowUpdate = mAdd
End If
Exit Sub
1:
MessageBox Err.Description, "frmcalendar:datagrid1_rowcolchange" & Err.Number, msgOkOnly, msgExclamation

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
GridLayout
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmCalendar
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT     CalendarID AS [Calendar ID], Description AS Keterangan, Day1, Day2, Day3, Day4, Day5, Day6, Day7, DateFrom AS [Dari Tanggal],                        DateTo AS [Sampai Tanggal] FROM         [Scheduling Calendar]"
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MyDDE.CheckRecordPendinged = True Then
   ScanKey vbKeyF5, 0, MyDDE
   If MyDDE.IsSucces = True Then
      Cancel = False
      MyDDE.ClearRecordset
      Set FrmCalendar = Nothing
   Else
      Cancel = True
   End If
Else
   MyDDE.ClearRecordset
   Set FrmCalendar = Nothing
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmCalendar = Nothing
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 1
Select Case AdReasonActiveDb
       Case tmbAddNew:
            mAdd = True
            txtBox(0).SetFocus
       Case tmbEdit:
            txtBox(0).Enabled = False
            mAdd = True
            txtBox(1).SetFocus
       Case tmbCancel:
            If MyDDE.ChildRecordset.Recordcount <> 0 Then
               mAdd = True
            Else
               mAdd = False
            End If
       Case tmbDetail:
            If MyDDE.IsChildMemberReady = True Then
               MyDDE.ChildRecordset.Fields("DateFrom") = Date
               MyDDE.ChildRecordset.Fields("DateTo") = Date
               MyDDE.ChildRecordset.Fields("Description") = "-"
            End If
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
                If MyDDE.ChildRecordset.Recordcount <> 0 Then
                   MyDDE.ChildRecordset.MoveFirst
                   If SendDataToServer("Delete From          [Scheduling Calendar Detail] WHERE     (CalendarID = N'" & txtBox(0) & "')") = True Then
                    Do
                    If MyDDE.ChildRecordset.EOF Then Exit Do
                       SendDataToServer " Insert Into [Scheduling Calendar Detail](CalendarID, DateFrom, DateTo, Description) " & _
                                        " values ('" & txtBox(0) & "',Convert(datetime,'" & Format(MyDDE.ChildRecordset.Fields("DateFrom"), "dd/mm/yy") & "',3),Convert(datetime,'" & Format(MyDDE.ChildRecordset.Fields("Dateto"), "dd/mm/yy") & "',3),N'" & MyDDE.ChildRecordset.Fields("Description") & "')"
                       MyDDE.ChildRecordset.MoveNext
                    Loop
                   End If
                   MyDDE.ChildRecordset.MoveLast
                End If
                mAdd = False
            End If
       Case tmbPrint:
            CallRPTReport "Calendar Report.rpt"
       Case Else: 'mVarDataDc = False
End Select
Exit Sub
1:
MessageBox Err.Description, "frmbommethode_mydde_afterpreparationdb" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail MyDDE.GetFieldByName("Calendar ID")
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 2
Dim mDel As New clsDelete
Select Case AdReasonActiveDb
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
'               If mDel.CekDelete(txtBox(0), reDelMasterCurency) = False Then
                  MyDDE.IsChildMemberReady = True
                  PrepareQuery
'               Else
'                  MyDDE.CancelTrans = True
'                  MessageBox "Record (" & txtBox(0) & ") Sedang Dipakai Transaksi Lain." & vbCrLf & "Record Tidak Bisa DiHapus.", "Peringatan", msgOkOnly
'                  MyDDE.IsChildMemberReady = False
'               End If
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
2:
MessageBox Err.Description, "frmcallendar_mydde_prepareactivedb" & Err.Number, msgOkOnly, msgExclamation

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
On Error GoTo xErr
With MyDDE
    .PrepareAppend = " INSERT INTO [Scheduling Calendar] ([CalendarID], [Description], Day1, Day2, Day3, Day4, Day5, Day6, Day7, DateFrom,DateTo) " & _
                     " VALUES (N'" & (txtBox(0)) & "', N'" & txtBox(1) & "'," & DaysCek(0) & "," & DaysCek(1) & "," & DaysCek(2) & "," & DaysCek(3) & "," & DaysCek(4) & "," & DaysCek(5) & "," & DaysCek(6) & ",Convert(Datetime,'" & Format(Tanggal(0), "dd/mm/yy") & "',3),Convert(Datetime,'" & Format(Tanggal(1), "dd/mm/yy") & "',3))"
                     
    .PrepareUpdate = " UPDATE [Scheduling Calendar] Set [Description] = N'" & txtBox(1) & "', day1=" & DaysCek(0) & ",day2= " & DaysCek(1) & ",day3=" & DaysCek(2) & ",day4=" & DaysCek(3) & ",day5=" & DaysCek(4) & ",day6=" & DaysCek(5) & ",day7=" & DaysCek(6) & ",datefrom = Convert(Datetime,'" & Format(Tanggal(0), "dd/mm/yy") & "',3),dateTo = Convert(Datetime,'" & Format(Tanggal(1), "dd/mm/yy") & "',3) WHERE     ([CalendarID] = N'" & txtBox(0) & "')"
    
    .PrepareDelete = " DELETE FROM [Scheduling Calendar] WHERE   ([CalendarID] = N'" & txtBox(0) & "') "
End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub GridLayout()
DataGrid1(0).Height = 2625
DataGrid1(0).width = 9390
DataGrid1(0).Columns(0).width = 1950.236
DataGrid1(0).Columns(1).width = 1950.236
DataGrid1(0).Columns(2).width = 4935.118
End Sub

Private Sub OpenDetail(ByVal Param As String)
On Error GoTo 1
Dim Rc As New DBQuick
Rc.DBOpen "SELECT     CalendarID, DateFrom, DateTo, Description FROM         [Scheduling Calendar Detail] WHERE     (CalendarID = N'" & Param & "') ORDER BY DateFrom", CNN, lckLockBatch
Set MyDDE.ChildRecordset = Rc.DBRecordset.Clone(adLockBatchOptimistic)
Set DataGrid1(0).DataSource = MyDDE.ChildRecordset
Exit Sub
1:
MessageBox Err.Description, "frmcalendar:opendetail" & Err.Number, msgOkOnly, msgExclamation
End Sub





