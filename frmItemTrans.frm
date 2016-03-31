VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmItemTrans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipe Item Transaksi"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9105
   Icon            =   "frmItemTrans.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   9105
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   0
      ScaleHeight     =   4695
      ScaleWidth      =   9105
      TabIndex        =   3
      Top             =   0
      Width           =   9105
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "keterangan"
         DataSource      =   "aDDE"
         Height          =   330
         Index           =   1
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   2
         Tag             =   "TP"
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "tipeid"
         DataSource      =   "aDDE"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   1680
         TabIndex        =   1
         Top             =   120
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid grid 
         Bindings        =   "frmItemTrans.frx":6852
         Height          =   3615
         Left            =   105
         TabIndex        =   4
         Tag             =   "TP"
         Top             =   960
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   6376
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "tipeid"
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
            DataField       =   "keterangan"
            Caption         =   "Keterangan"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
            EndProperty
         EndProperty
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   1680
         X2              =   240
         Y1              =   1120
         Y2              =   1120
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   1680
         X2              =   240
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1680
         X2              =   240
         Y1              =   435
         Y2              =   435
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
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   525
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
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
         TabIndex        =   5
         Top             =   165
         Width           =   615
      End
   End
   Begin SemeruDC.SemeruOleDC aDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   4665
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   1005
      BindFormTAG     =   "TP"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmItemTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub aDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
   Select Case AdReasonActiveDb
      Case tmbAddNew, tmbEdit:
         txt(1).SetFocus
         txt(0).Enabled = False
      Case tmbSave:
         aDDE.RefreshDatabase
   End Select

End Sub

Private Sub aDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   txt(0).Text = aDDE.GetFieldByName("tipeid")
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
   Case tmbDelete:
      PrepareQuery
End Select
End Sub

Private Sub PrepareQuery()
On Error GoTo xErr
   With aDDE
      .PrepareAppend = "insert into tipe_itemtrans (Keterangan) values('" & .GetFieldByName("keterangan") & "')"
                                        
      .PrepareUpdate = "update tipe_itemtrans set keterangan = '" & .GetFieldByName("keterangan") & "' where tipeid=" & .GetFieldByName("tipeid")
                                            
      .PrepareDelete = "delete from tipe_itemtrans where tipeid = " & .GetFieldByName("tipeid")
      
   End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub Form_Load()
   'HiasForm Picture1, Me
   HiasFormManTell Picture2, Me
   aDDE.SetPermissions = aksess.MayDo("Tipe Item Transaksi") 'set hak aksess
   Set aDDE.BindForm = Me
   Set aDDE.ActiveConnection = CNN
   aDDE.PrepareQuery = "select * from tipe_itemtrans"
   Set grid.DataSource = aDDE.ActiveRecordset
End Sub


