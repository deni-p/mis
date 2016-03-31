VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmMetodeBayar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Metode Pembayaran"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMotodeBayar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   9855
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
      Height          =   4845
      Left            =   0
      ScaleHeight     =   4845
      ScaleWidth      =   9855
      TabIndex        =   5
      Top             =   0
      Width           =   9855
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "code"
         DataSource      =   "aDDE"
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   1
         Tag             =   "TP"
         Top             =   120
         Width           =   1815
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "Bal_ Account Type"
         DataSource      =   "aDDE"
         Height          =   315
         Index           =   1
         Left            =   1680
         TabIndex        =   2
         Tag             =   "TP"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "Bal_ Account No_"
         DataSource      =   "aDDE"
         Height          =   315
         Index           =   2
         Left            =   1680
         TabIndex        =   3
         Tag             =   "TP"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "description"
         DataSource      =   "aDDE"
         Height          =   1035
         Index           =   3
         Left            =   5400
         MultiLine       =   -1  'True
         TabIndex        =   4
         Tag             =   "TP"
         Top             =   120
         Width           =   4335
      End
      Begin MSDataGridLib.DataGrid grid 
         Bindings        =   "frmMotodeBayar.frx":6852
         Height          =   3255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   5741
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "Code"
            Caption         =   "Kode"
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
            DataField       =   "Bal_ Account Type"
            Caption         =   "Bal Account Type"
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
         BeginProperty Column02 
            DataField       =   "Bal_ Account No_"
            Caption         =   "Bal Account No"
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
         BeginProperty Column03 
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
            BeginProperty Column00 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bal Account Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bal Account No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   4260
         TabIndex        =   7
         Top             =   120
         Width           =   1035
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1680
         X2              =   240
         Y1              =   400
         Y2              =   400
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   1680
         X2              =   240
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   1680
         X2              =   240
         Y1              =   1120
         Y2              =   1120
      End
   End
   Begin SemeruDC.SemeruOleDC aDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   4845
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   1005
      BindFormTAG     =   "TP"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmMetodeBayar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub aDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
   Select Case AdReasonActiveDb
      Case tmbAddNew:
         txt(0).SetFocus
         aDDE.GetFieldByName("Bal_ Account Type") = 0
         aDDE.GetFieldByName("Bal_ Account No_") = " "
         aDDE.GetFieldByName("description") = "-"
         
   End Select

End Sub

Private Sub aDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
   Case tmbSave:
      If aDDE.CheckEmptyControl = False Then
         aDDE.IsChildMemberReady = True
         PrepareQuery
      Else
         aDDE.IsChildMemberReady = False
      End If
End Select
End Sub

Private Sub PrepareQuery()
   With aDDE
      .PrepareAppend = "insert into TermMethod (code,Description,[Bal_ Account Type],[Bal_ Account No_]) values('" & _
                                                .GetFieldByName("code") & _
                                        "','" & .GetFieldByName("description") & _
                                        "', " & .GetFieldByName("Bal_ Account Type") & _
                                        ", '" & .GetFieldByName("Bal_ Account No_") & "')"
                                        
      .PrepareUpdate = "update TermMethod set [Bal_ Account Type] = " & .GetFieldByName("Bal_ Account Type") & _
                                            ", [Bal_ Account No_]='" & .GetFieldByName("Bal_ Account No_") & _
                                            "', description ='" & .GetFieldByName("description") & "' where code ='" & .GetFieldByName("code") & "'"
                                            
      .PrepareDelete = "delete from TermMethod where code = '" & .GetFieldByName("code") & "'"
   End With
End Sub

Private Sub Form_Load()
   'HiasForm Picture1, Me
   HiasFormManTell Picture2, Me
   Set aDDE.BindForm = Me
   Set aDDE.ActiveConnection = CNN
   aDDE.PrepareQuery = "select * from TermMethod"
   Set grid.DataSource = aDDE.ActiveRecordset
End Sub

