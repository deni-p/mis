VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{82968C93-C596-4A47-8A14-646737648F29}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmQuarter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quarter"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   9975
   Tag             =   "Quarter Setting"
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2835
      Left            =   45
      ScaleHeight     =   2805
      ScaleWidth      =   9855
      TabIndex        =   10
      Top             =   120
      Width           =   9885
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   2190
         Left            =   120
         ScaleHeight     =   2160
         ScaleWidth      =   9525
         TabIndex        =   11
         Top             =   360
         Width           =   9555
         Begin VB.TextBox Text1 
            BorderStyle     =   0  'None
            DataField       =   "Year"
            Height          =   315
            Left            =   1545
            MaxLength       =   4
            TabIndex        =   1
            Tag             =   "ASM"
            Text            =   "2005"
            Top             =   120
            Width           =   930
         End
         Begin VB.ComboBox Combo1 
            DataField       =   "Quarter"
            Height          =   315
            ItemData        =   "FrmQuarter.frx":0000
            Left            =   4590
            List            =   "FrmQuarter.frx":0025
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Tag             =   "ASM"
            Top             =   120
            Width           =   840
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "StartDate"
            Height          =   315
            Index           =   1
            Left            =   1545
            TabIndex        =   5
            Tag             =   "ASM"
            Top             =   780
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MMM/yyyy"
            Format          =   57147395
            CurrentDate     =   38272
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "StopDate"
            Height          =   315
            Index           =   2
            Left            =   1545
            TabIndex        =   7
            Tag             =   "ASM"
            Top             =   1110
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MMM/yyyy"
            Format          =   57147395
            CurrentDate     =   38272
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "MidDate"
            Height          =   315
            Index           =   3
            Left            =   1545
            TabIndex        =   9
            Tag             =   "ASM"
            Top             =   1440
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MMM/yyyy"
            Format          =   57147395
            CurrentDate     =   38272
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "IDX"
            DataField       =   "IDX"
            Height          =   195
            Left            =   4680
            TabIndex        =   13
            Tag             =   "ASM"
            Top             =   1350
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            X1              =   375
            X2              =   9360
            Y1              =   2010
            Y2              =   2010
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   3750
            X2              =   5385
            Y1              =   420
            Y2              =   405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quarter                          Nd"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   4
            Left            =   3750
            TabIndex        =   2
            Top             =   165
            Width           =   1935
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   375
            X2              =   2205
            Y1              =   1740
            Y2              =   1740
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   375
            X2              =   2205
            Y1              =   1395
            Y2              =   1395
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mid Date"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   3
            Left            =   375
            TabIndex        =   8
            Top             =   1500
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End Date"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   2
            Left            =   375
            TabIndex        =   6
            Top             =   1170
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Year"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   0
            Left            =   375
            TabIndex        =   0
            Top             =   180
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   1
            Left            =   375
            TabIndex        =   4
            Top             =   840
            Width           =   750
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   375
            X2              =   2205
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   375
            X2              =   2205
            Y1              =   1080
            Y2              =   1080
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   12
      Top             =   3000
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmQuarter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DTPicker1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Form_Load()
HiasForm Picture1, Me
CenterForm Picture2, Me
DTPicker1(1).Value = Date
DTPicker1(2).Value = Date
DTPicker1(3).Value = Date
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmQuarter
    .BindFormTAG = "ASM"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT  * From SetupQuarterData"
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
End Sub

Private Sub Form_Resize()
'
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmQuarter = Nothing
End Sub



Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
MyDDE.PrepareAppend = " INSERT INTO SetupQuarterData" & _
                      " (Idx, [Year], Quarter, StartDate, StopDate, MidDate)" & _
                      " VALUES (NEWID(), " & Text1 & ", " & IIf(Combo1.Text <> "", Combo1.Text, 1) & " , CONVERT(DATETIME, '" & Format(DTPicker1(1).Value, "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format(DTPicker1(2).Value, "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format(DTPicker1(3).Value, "dd/mm/yy") & "', 3))"
MyDDE.PrepareUpdate = " UPDATE SetupQuarterData " & _
                      " SET [Year] = " & Text1 & ", Quarter =" & IIf(Combo1.Text <> "", Combo1.Text, 1) & ", StartDate = CONVERT(DATETIME, '" & Format(DTPicker1(1).Value, "dd/mm/yy") & "', 3), StopDate = CONVERT(DATETIME, '" & Format(DTPicker1(3).Value, "dd/mm/yy") & "', 3)," & _
                      " MidDate = CONVERT(DATETIME, '" & Format(DTPicker1(3).Value, "dd/mm/yy") & "', 3) WHERE     (Idx = '" & Label2 & "')"
MyDDE.PrepareDelete = " DELETE FROM SetupQuarterData WHERE     (Idx = '" & Label2 & "')"
Err.Clear
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbAddNew:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               Text1 = Format(Year(Date), "000#")
               Text1.SetFocus
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbEdit:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               Text1.SetFocus
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
'               If mDel.CekDelete(txtBox(0), reDelMasterCurency) = False Then
                  MyDDE.IsChildMemberReady = True
''                  PrepareQuery
'               Else
'                  MyDDE.CancelTrans = True
''                  MessageBox "Record (" & txtBox(0) & ") Sedang Dipakai Transaksi Lain." & vbCrLf & "Record Tidak Bisa DiHapus.", "Peringatan", msgOkOnly
'                  MyDDE.IsChildMemberReady = False
'               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
'               PrepareQuery
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Picture1.Parent.hwnd
End Sub

Private Sub Text1_Change()
If Text1 = "" Then Text1 = Format(Year(Date), "000#")
End Sub
