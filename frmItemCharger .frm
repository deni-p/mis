VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmItemCharge 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detail Jasa"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11190
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmItemCharger .frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   11190
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
      Height          =   5160
      Left            =   0
      ScaleHeight     =   5160
      ScaleWidth      =   11190
      TabIndex        =   10
      Top             =   0
      Width           =   11190
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   9960
         Picture         =   "frmItemCharger .frx":6852
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   855
         Width           =   330
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "No_"
         DataSource      =   "aDDE"
         Height          =   330
         Index           =   0
         Left            =   2520
         TabIndex        =   1
         Tag             =   "TP"
         Top             =   120
         Width           =   2205
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "Description"
         DataSource      =   "aDDE"
         Height          =   330
         Index           =   1
         Left            =   2520
         TabIndex        =   2
         Tag             =   "TP"
         Top             =   480
         Width           =   2205
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "Gen_ Prod_ Posting Group"
         DataSource      =   "aDDE"
         Height          =   330
         Index           =   2
         Left            =   2520
         TabIndex        =   3
         Tag             =   "TP"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "Tax Group Code"
         DataSource      =   "aDDE"
         Height          =   330
         Index           =   3
         Left            =   2520
         TabIndex        =   4
         Tag             =   "TP"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "VAT Prod_ Posting Group"
         DataSource      =   "aDDE"
         Height          =   330
         Index           =   4
         Left            =   7665
         MultiLine       =   -1  'True
         TabIndex        =   5
         Tag             =   "TP"
         Top             =   120
         Width           =   2895
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "Search Description"
         DataSource      =   "aDDE"
         Height          =   330
         Index           =   5
         Left            =   7665
         MultiLine       =   -1  'True
         TabIndex        =   6
         Tag             =   "TP"
         Top             =   480
         Width           =   2640
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "Global Dimension 1 Code"
         DataSource      =   "aDDE"
         Height          =   330
         Index           =   6
         Left            =   7665
         MultiLine       =   -1  'True
         TabIndex        =   7
         Tag             =   "TP"
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "Global Dimension 2 Code"
         DataSource      =   "aDDE"
         Height          =   330
         Index           =   7
         Left            =   7665
         MultiLine       =   -1  'True
         TabIndex        =   8
         Tag             =   "TP"
         Top             =   1200
         Width           =   2295
      End
      Begin MSDataGridLib.DataGrid grid 
         Bindings        =   "frmItemCharger .frx":6BDC
         Height          =   3015
         Left            =   120
         TabIndex        =   9
         Tag             =   "TP"
         Top             =   1800
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   5318
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   3
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "no_"
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
            DataField       =   "Description"
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
         BeginProperty Column02 
            DataField       =   "[Gen_ Prod_ Posting Group]"
            Caption         =   "Gen Prod Posting Group"
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
            DataField       =   "Tax Group Code"
            Caption         =   "Tax Group Code"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "VAT Prod_ Posting Group"
            Caption         =   "VAT Prod Posting Group"
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
            DataField       =   "Search Description"
            Caption         =   "Search Description"
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
            DataField       =   "Global Dimension 1 Code"
            Caption         =   "Global Dimension 1 Code"
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
            DataField       =   "Global Dimension 2 Code"
            Caption         =   "Global Dimension 2 Code"
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
            BeginProperty Column04 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
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
         TabIndex        =   18
         Top             =   158
         Width           =   615
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
         TabIndex        =   17
         Top             =   518
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Gen  Prod Posting Group"
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
         TabIndex        =   16
         Top             =   878
         Width           =   3150
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Group Code"
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
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   1238
         Width           =   2370
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "VAT Prod Posting Group"
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
         Left            =   5280
         TabIndex        =   14
         Top             =   158
         Width           =   2340
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Description"
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
         Index           =   5
         Left            =   5280
         TabIndex        =   13
         Top             =   518
         Width           =   2760
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Global Dimension 1 Code"
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
         Index           =   6
         Left            =   5280
         TabIndex        =   12
         Top             =   878
         Width           =   2355
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Global Dimension 2 Code"
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
         Index           =   7
         Left            =   5280
         TabIndex        =   11
         Top             =   1238
         Width           =   2445
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2520
         X2              =   240
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2520
         X2              =   240
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   2520
         X2              =   240
         Y1              =   1155
         Y2              =   1155
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   2520
         X2              =   240
         Y1              =   1515
         Y2              =   1515
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   7920
         X2              =   5280
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   7890
         X2              =   5280
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   7860
         X2              =   5280
         Y1              =   1155
         Y2              =   1155
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   7845
         X2              =   5280
         Y1              =   1515
         Y2              =   1515
      End
   End
   Begin SemeruDC.SemeruOleDC aDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5145
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   1005
      BindFormTAG     =   "TP"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmItemCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RsLookup As New DBQuick

Private Sub aDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
   cmdLink(0).Enabled = False
   Select Case AdReasonActiveDb
      Case tmbAddNew:
         cmdLink(0).Enabled = True
         txt(0).SetFocus
         With aDDE
          .GetFieldByName("No_") = ""
          .GetFieldByName("Description") = ""
          .GetFieldByName("Gen_ Prod_ Posting Group") = ""
          .GetFieldByName("Tax Group Code") = ""
          .GetFieldByName("VAT Prod_ Posting Group") = ""
          .GetFieldByName("Search Description") = ""
          .GetFieldByName("Global Dimension 1 Code") = ""
          .GetFieldByName("Global Dimension 2 Code") = ""
         End With
      Case tmbEdit:
         cmdLink(0).Enabled = True
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
On Error GoTo xErr
'   [timestamp] [binary] (8) NULL ,
'   [No_] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
'   [Description] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
'   [Gen_ Prod_ Posting Group] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
'   [Tax Group Code] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
'   [VAT Prod_ Posting Group] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
'   [Search Description] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
'   [Global Dimension 1 Code] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
'   [Global Dimension 2 Code] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL

   With aDDE
      .PrepareAppend = "insert into item_charge ([No_]," & _
                                               " [Description]," & _
                                               " [Gen_ Prod_ Posting Group], " & _
                                               " [Tax Group Code] ," & _
                                               " [VAT Prod_ Posting Group] ," & _
                                               " [Search Description]," & _
                                               " [Global Dimension 1 Code]," & _
                                               " [Global Dimension 2 Code]) " & _
                        " values ('" & .GetFieldByName("No_") & _
                               "','" & .GetFieldByName("Description") & _
                               "','" & .GetFieldByName("Gen_ Prod_ Posting Group") & _
                               "','" & .GetFieldByName("Tax Group Code") & _
                               "','" & .GetFieldByName("VAT Prod_ Posting Group") & _
                               "','" & .GetFieldByName("Search Description") & _
                               "','" & .GetFieldByName("Global Dimension 1 Code") & _
                               "','" & .GetFieldByName("Global Dimension 2 Code") & "')"
                               
                               
      .PrepareUpdate = "update item_charge set [Description] = '" & .GetFieldByName("Description") & _
                                            "', [Gen_ Prod_ Posting Group]='" & .GetFieldByName("Gen_ Prod_ Posting Group") & _
                                            "', [Tax Group Code]='" & .GetFieldByName("Tax Group Code") & _
                                            "', [VAT Prod_ Posting Group] ='" & .GetFieldByName("VAT Prod_ Posting Group") & _
                                            "', [Search Description]='" & .GetFieldByName("Search Description") & _
                                            "', [Global Dimension 1 Code]='" & .GetFieldByName("Global Dimension 1 Code") & _
                                            "', [Global Dimension 2 Code] ='" & .GetFieldByName("Global Dimension 2 Code") & _
                       "' where No_ ='" & .GetFieldByName("No_") & "'"
                       
      .PrepareDelete = "delete from item_charge where No_ = '" & .GetFieldByName("No_") & "'"
   End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub cmdLink_Click(Index As Integer)
   RsLookup.DBOpen "Select noItem as Kode,ItemName as [Nama Barang],UOM as Satuan from inventory where categID='JS' and noGroup='CH'", CNN
   Set mCall.FormData = RsLookup.DBRecordset
   mCall.FromTagActive = "Inventory"
End Sub

Private Sub Form_Load()
   'HiasForm Picture1, Me
   HiasFormManTell Picture2, Me
   aDDE.SetPermissions = aksess.MayDo("Detail Item Jasa") 'set hak aksess

   Set aDDE.BindForm = Me
   Set aDDE.ActiveConnection = CNN
   aDDE.PrepareQuery = "select * from Item_charge"
   Set grid.DataSource = aDDE.ActiveRecordset
   grid.HeadLines = 3
   Set mCall = New frmCaller
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
   aDDE.GetFieldByName("Global Dimension 1 Code") = mCall.GetFieldByName(0)
End Sub
