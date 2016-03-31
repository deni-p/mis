VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmUserAktive 
   AutoRedraw      =   -1  'True
   Caption         =   "User Connection"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7530
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmUserAktive.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6060
   ScaleWidth      =   7530
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5760
      Left            =   45
      ScaleHeight     =   5700
      ScaleWidth      =   9300
      TabIndex        =   0
      Top             =   240
      Width           =   9360
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000010&
         Height          =   5100
         Left            =   90
         ScaleHeight     =   5040
         ScaleWidth      =   8205
         TabIndex        =   1
         Top             =   375
         Width           =   8265
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   4635
            Left            =   90
            TabIndex        =   2
            Top             =   165
            Width           =   8040
            _ExtentX        =   14182
            _ExtentY        =   8176
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BorderStyle     =   0
            HeadLines       =   2
            RowHeight       =   15
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
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   8
            BeginProperty Column00 
               DataField       =   "spid"
               Caption         =   "Spid"
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
               DataField       =   "ecid"
               Caption         =   "Ecid"
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
               DataField       =   "status"
               Caption         =   "Status"
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
               DataField       =   "loginame"
               Caption         =   "Login Name"
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
               DataField       =   "hostname"
               Caption         =   "Host Name"
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
               DataField       =   "blk"
               Caption         =   "BLK"
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
            BeginProperty Column06 
               DataField       =   "dbname"
               Caption         =   "Database Name"
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
            BeginProperty Column07 
               DataField       =   "cmd"
               Caption         =   "Cmd"
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
                  ColumnWidth     =   615.118
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   615.118
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   764.787
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   1590.236
               EndProperty
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "FrmUserAktive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rc As Recordset

Private Sub Form_Load()

Set DataGrid1.DataSource = OpenUserAccess
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState <> vbMaximized Then
   Me.Height = MainMenu.ScaleHeight
   Me.Width = MainMenu.ScaleWidth
End If
HiasForm Picture1, Me
CenterForm Picture2
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
End Sub
