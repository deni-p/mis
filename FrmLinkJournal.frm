VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FrmLinkJournal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seting Value Journal"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmLinkJournal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1635
      TabIndex        =   7
      Top             =   4950
      Width           =   1395
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5460
      Left            =   90
      ScaleHeight     =   5400
      ScaleWidth      =   7935
      TabIndex        =   4
      Top             =   105
      Width           =   7995
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Kredit"
         Height          =   300
         Left            =   450
         TabIndex        =   1
         Top             =   3615
         Width           =   1305
      End
      Begin MSDataGridLib.DataGrid DgJournal 
         Height          =   3015
         Left            =   135
         TabIndex        =   0
         Top             =   585
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   5318
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox TxtValue 
         Height          =   315
         Index           =   1
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   4290
         Width           =   3570
      End
      Begin VB.TextBox TxtValue 
         Height          =   315
         Index           =   0
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   3945
         Width           =   3570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "List Value"
         Height          =   210
         Index           =   1
         Left            =   450
         TabIndex        =   6
         Top             =   4335
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "List Record"
         Height          =   210
         Index           =   0
         Left            =   450
         TabIndex        =   5
         Top             =   4005
         Width           =   900
      End
   End
End
Attribute VB_Name = "FrmLinkJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RcJournal As New DBQuick
Private mVarGroupJournal As String
Private mValst As New Recordset

Private Sub Check1_Click()
If Check1.Value = 1 Then
   mVarSetingJournalData.Posisijournal = 1
   Check1.Caption = "Debet"
Else
   Check1.Caption = "Kredit"
   mVarSetingJournalData.Posisijournal = 0
End If
End Sub

Private Sub CmdOK_Click()
Unload Me
End Sub

Private Sub DataGrid1_Error(ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub

Private Sub DgJournal_ButtonClick(ByVal ColIndex As Integer)
Dim I As Integer
If TxtValue(0) = "" Or TxtValue(1) = "" Then
   If TxtValue(0) = "" Then
      TxtValue(0) = mValst.Fields(0).Value
      mVarSetingJournalData.JournalFieldString = mValst.Fields(0).Value
   ElseIf TxtValue(1) = "" Then
      TxtValue(1) = mValst.Fields(1).Value
      If IsNumeric(TxtValue(1)) = True Then
         mVarSetingJournalData.JournalValueString = mValst.Fields(0).Value
      Else
         MessageBox "Data harus berisi numeric.Harap diulangi", "Peringatan", msgOkOnly
         TxtValue(1) = ""
      End If
   End If
Else
   I = MessageBox("Tidak boleh mengisi variabel seting journal lagi.Variabel harus dibersihkan lagi jika ingin membuat setup." & vbCrLf & "Tekan YES untuk membersihkan Setup", "Peringatan", msgYesNo)
   If I = 1 Then
      TxtValue(0) = ""
      TxtValue(1) = ""
   End If
End If
End Sub

Private Sub DgJournal_Error(ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub

Private Sub Form_Load()
mVarSetingJournalData.Posisijournal = 0
JournalValueString = ""
If RcJournal.DBOpen(" Select * From [Journal " & mVarGroupJournal & "]", Cnn, lckLockReadOnly) = True Then
'Set DgJournal.DataSource = RcJournal.DBRecordset
ListFieldData
RcJournal.CloseDB
Else
   MessageBox "Data Tabel Journal belum ada.", "Peringatan", msgOkOnly
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next

mValst.Close
'mValFldt.Close
Err.Clear
End Sub

Private Sub Form_Resize()
On Error Resume Next
HiasForm Picture1, Me
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
End Sub

Public Property Get GroupJournal() As String
GroupJournal = mVarGroupJournal
End Property

Public Property Let GroupJournal(ByVal vNewValue As String)
mVarGroupJournal = vNewValue
End Property

'Private Sub TotoGrid()
'Dim I As Integer
'With RcJournal.DBRecordset
'     For I = 0 To .Fields.Count - 1
'        Select Case .Fields(I).Type
'               Case adBigInt, adCurrency, adDecimal, adDouble, adInteger
'                    Mystd.Type = fmtCustom
'                    Mystd.Format = "#,##0"
'                    Set DgJournal.Columns(I).DataFormat = Mystd
'                    DgJournal.Columns(I).Alignment = dbgRight
'                    DgJournal.Columns(I).Button = True
'               Case adDate, adDBDate, adDBTime, adDBTimeStamp
'                    Mystd.Type = fmtCustom
'                    Mystd.Format = "dd/MM/yyyy"
'                    'Set DgJournal.Columns(I).DataFormat = myStd
'                    DgJournal.Columns(I).Alignment = dbgRight
'               Case Else: DgJournal.Columns(I).Button = False
'
'        End Select
'     Next I
'End With
'End Sub

Private Sub ListFieldData()
Dim I As Integer
'mValFldt.Fields.Append "Tag Record", adBSTR
'mValFldt.Open

mValst.Fields.Append "List Field", adBSTR
mValst.Fields.Append "Sample Data", adBSTR
mValst.Open
With RcJournal.DBRecordset
     For I = 0 To .Fields.Count - 1
         
         If .Recordcount <> 0 Then
             Select Case .Fields(I).Type
                    Case adBigInt, adCurrency, adDecimal, adDouble, adInteger
                         mValst.AddNew 0, .Fields(I).Name
                         mValst.Fields(1) = FormatNumber(IIf(Not IsNull(.Fields(I).Value), .Fields(I), 0), 0)
                    Case adDate, adDBDate, adDBTime, adDBTimeStamp
                         'mValst.Fields(1) = Format(IIf(Not IsNull(.Fields(I).Value), .Fields(I), Date), "dd/mmmm/yyyy")
                    Case 11:
                    Case Else:
                         mValst.AddNew 0, .Fields(I).Name
                         mValst.Fields(1) = IIf(Not IsNull(.Fields(I).Value), .Fields(I), "-")
             End Select
          Else
             Select Case .Fields(I).Type
                    Case adBigInt, adCurrency, adDecimal, adDouble, adInteger
                         mValst.AddNew 0, .Fields(I).Name
                         mValst.Fields(1) = Rnd(10000)
                    Case adDate, adDBDate, adDBTime, adDBTimeStamp
                         mValst.AddNew 0, .Fields(I).Name
                         mValst.Fields(1) = Format(Date, "dd/mm/yyyy")
                    Case 11:
                    Case Else:
                         mValst.AddNew 0, .Fields(I).Name
                         mValst.Fields(1) = "xxx"
             End Select
          End If
     Next I
End With
Set DgJournal.DataSource = mValst
DgJournal.Columns(0).Width = 3500
DgJournal.Columns(1).Width = 3500
DgJournal.Columns(0).Alignment = dbgLeft
DgJournal.Columns(1).Alignment = dbgRight
DgJournal.AllowUpdate = False
DgJournal.Columns(0).Button = True
DgJournal.Refresh
End Sub
