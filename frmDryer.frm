VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmDryer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dryer"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   10605
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDryer.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   10605
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5385
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   1005
      BindFormTAG     =   "dryer"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
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
      Height          =   5370
      Left            =   0
      ScaleHeight     =   5370
      ScaleWidth      =   10605
      TabIndex        =   14
      Top             =   0
      Width           =   10605
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "noekstraksi"
         DataSource      =   "DDE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   4230
         TabIndex        =   4
         Tag             =   "dryer"
         Top             =   525
         Width           =   1980
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   5715
         Picture         =   "frmDryer.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   525
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   6225
         Picture         =   "frmDryer.frx":6BDC
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   173
         Visible         =   0   'False
         Width           =   390
      End
      Begin MSComCtl2.DTPicker viewDate 
         Height          =   405
         Left            =   3285
         TabIndex        =   12
         Top             =   3045
         Visible         =   0   'False
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   714
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy  HH:mm"
         Format          =   59375619
         CurrentDate     =   39534
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   2910
         Left            =   60
         TabIndex        =   10
         Tag             =   "dryer"
         Top             =   2280
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   5133
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   1
      End
      Begin MSDataGridLib.DataGrid GridDry 
         Height          =   2625
         Left            =   2280
         TabIndex        =   11
         Top             =   2580
         Width           =   8250
         _ExtentX        =   14552
         _ExtentY        =   4630
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   2
         RowHeight       =   15
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "dryer"
            Caption         =   "Dryer"
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
            DataField       =   "no"
            Caption         =   "Urutan"
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
            DataField       =   "tanggal_mulai"
            Caption         =   "Tgl Mulai"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd MMM yyyy  hh:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "tanggal_selesai"
            Caption         =   "Tgl Selesai"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd MMM yyyy  hh:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "QtyBasah"
            Caption         =   "Qty Basah"
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
         BeginProperty Column05 
            DataField       =   "QtyKering"
            Caption         =   "Qty Kering"
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
         BeginProperty Column06 
            DataField       =   "rendemen"
            Caption         =   "Rendemen"
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
         BeginProperty Column07 
            DataField       =   "kekeringan"
            Caption         =   "Kekeringan"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Kering"
               FalseValue      =   "Basah"
               NullValue       =   "Basah"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
               Button          =   -1  'True
               Locked          =   -1  'True
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "keterangan"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   780
         Index           =   6
         Left            =   5835
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Tag             =   "dryer"
         Top             =   1200
         Width           =   4605
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "grup"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   8775
         TabIndex        =   6
         Tag             =   "dryer"
         Top             =   525
         Width           =   1650
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "OrderID"
         DataSource      =   "DDE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   4230
         Locked          =   -1  'True
         TabIndex        =   2
         Tag             =   "dryer"
         Top             =   165
         Width           =   1995
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "id"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   1
         Tag             =   "dryer"
         Top             =   165
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "tanggal_mulai"
         DataSource      =   "dde"
         Height          =   315
         Index           =   0
         Left            =   2205
         TabIndex        =   7
         Tag             =   "dryer"
         Top             =   1215
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy    HH:mm"
         Format          =   59375619
         CurrentDate     =   39419
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "tanggal"
         DataSource      =   "dde"
         Height          =   315
         Index           =   2
         Left            =   8775
         TabIndex        =   5
         Tag             =   "dryer"
         Top             =   165
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   59375619
         CurrentDate     =   39365
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "tanggal_selesai"
         DataSource      =   "dde"
         Height          =   315
         Index           =   1
         Left            =   2220
         TabIndex        =   8
         Tag             =   "dryer"
         Top             =   1575
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy    HH:mm"
         Format          =   59375619
         CurrentDate     =   39419
      End
      Begin VB.Label lblDryer 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2280
         TabIndex        =   23
         Top             =   2280
         Width           =   8250
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   4695
         X2              =   2655
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "No Ekstraksi"
         Height          =   255
         Index           =   5
         Left            =   2685
         TabIndex        =   22
         Top             =   570
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan :"
         Height          =   255
         Index           =   33
         Left            =   5850
         TabIndex        =   20
         Top             =   990
         Width           =   1050
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   9300
         X2              =   7260
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   9300
         X2              =   7260
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal && waktu selesai"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   19
         Top             =   1620
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal && waktu mulai"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   1260
         Width           =   1890
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   2040
         X2              =   240
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   255
         Index           =   12
         Left            =   285
         TabIndex        =   17
         Top             =   180
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Manufacture Order"
         Height          =   255
         Index           =   0
         Left            =   2685
         TabIndex        =   16
         Top             =   210
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         Height          =   255
         Index           =   2
         Left            =   7275
         TabIndex        =   15
         Top             =   585
         Width           =   2055
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   4695
         X2              =   2655
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2235
         X2              =   195
         Y1              =   1515
         Y2              =   1515
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   4050
         X2              =   195
         Y1              =   1875
         Y2              =   1875
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   255
         Index           =   1
         Left            =   7290
         TabIndex        =   21
         Top             =   210
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmDryer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RsDetail As New DBQuick
Private MEdit As Boolean
Private IDGen As New IDGenerator
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RsLookup As New DBQuick
Private lMode As Byte
Private currCol As Integer

Public Property Let SetMode(ByVal Value As Byte)
   lMode = Value
End Property


Private Sub loadDetail()
    RsDetail.DBOpen "select dryer,no,tanggal_mulai,tanggal_selesai,qtyKering,qtyBasah,rendemen,kekeringan from dryer_detail where id = '" & MyDDE.GetFieldByName("id") & "'", CNN
    Set MyDDE.ChildRecordset = RsDetail.DBRecordset
    Set GridDry.DataSource = MyDDE.ChildRecordset
End Sub

Private Sub LoadDryer()
On Error GoTo 1
   Dim rsDryer As New DBQuick
   Dim vNode As Node

   
   rsDryer.DBOpen "select * from dryer", CNN
   If rsDryer.DBRecordset.Recordcount > 0 Then

      While Not rsDryer.DBRecordset.EOF
         Set vNode = TreeView1.Nodes.Add(, , rsDryer.DBRecordset.Fields("ID"), rsDryer.DBRecordset.Fields("ID"))
         rsDryer.DBRecordset.MoveNext
      Wend
      TreeView1.Nodes.Item(1).Selected = True
      'TreeView1_NodeClick TreeView1.SelectedItem
   End If
Exit Sub
1:
MessageBox Err.Description, "frmdryer:loaddryer" & Err.Number, msgOkOnly, msgExclamation
End Sub




Private Sub cmdLink_Click(Index As Integer)
On Error GoTo 1
   Select Case Index
      Case 0: RsLookup.DBOpen "select OrderID,OrderNAme,Type,RequireDate from [Manufacture Order] where status='RELEASED'", CNN
      Case 1: RsLookup.DBOpen "select NoEkstraksi from statusproduksi where (posisi = 'JEMUR' or posisi='CUTTER') and status=1", CNN
   End Select
   
   If RsLookup.DBRecordset.Recordcount > 0 Then
      Set mCall.FormData = RsLookup.DBRecordset
      Select Case Index
         Case 0: mCall.FromTagActive = "Manufacture Order"
         Case 1: mCall.FromTagActive = "No Ekstraksi"
      End Select
   Else
      MessageBox "Data Tidak Tersedia", "Stop", msgOkOnly, msgCrtical
   End If
Exit Sub
1:
MessageBox Err.Description, "frmdryer:cmdlink_click" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub GridDry_AfterColEdit(ByVal ColIndex As Integer)
   On Error GoTo xErr
   Select Case ColIndex
      Case 4, 5: GridDry.Columns(6).Value = Val(GridDry.Columns(5).Value) / Val(GridDry.Columns(4).Value) * 100
   End Select
Exit Sub
xErr:
   'MessageBox Err.Description, "Error", msgOkOnly, msgExclamation
   Err.Clear
End Sub

Private Sub GridDry_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo 1
   If GridDry.Columns(ColIndex).Value = True Then
      GridDry.Columns(ColIndex).Value = False
   Else
      GridDry.Columns(ColIndex).Value = True
   End If
Exit Sub
1:
MessageBox Err.Description, "frmdryer:gridry_buttonclick" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub gridDry_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo Hell
   currCol = 0
   If MyDDE.ChildRecordset.Recordcount > 0 Then
      viewDate.Visible = False
      Select Case GridDry.col
         Case 2: viewDate.Visible = True
                 viewDate.Value = MyDDE.ChildRecordset.Fields("tanggal_mulai")
                 viewDate.Move GridDry.Left + GridDry.Columns(2).Left, _
                               GridDry.Top + GridDry.RowTop(GridDry.row), _
                               GridDry.Columns(2).width, _
                               GridDry.RowHeight
                             
         Case 3: viewDate.Visible = True
                 viewDate.Value = MyDDE.ChildRecordset.Fields("tanggal_selesai")
                 viewDate.Move GridDry.Left + GridDry.Columns(3).Left, _
                               GridDry.Top + GridDry.RowTop(GridDry.row), _
                               GridDry.Columns(3).width, _
                               GridDry.RowHeight
      End Select
      currCol = GridDry.col
   End If
Exit Sub
Hell:
   If Err.Number = 380 Then
      viewDate.Value = Now
      Err.Clear
   End If
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
   Select Case UCase(mCall.FromTagActive)
      Case "MANUFACTURE ORDER": txt(1).Text = mCall.GetFieldByName(0)
      Case "NO EKSTRAKSI": txt(3).Text = mCall.GetFieldByName(0)
   End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 1
Dim x As Integer
Select Case AdReasonActiveDb
    Case tmbAddNew
       txt(0).Text = IDGen.GetID("DR")
       MyDDE.GetFieldByName("keterangan") = "-"
       MyDDE.GetFieldByName("noEkstraksi") = frmProduksi.txtBox(5).Text
       MyDDE.GetFieldByName("tanggal") = Now
       MyDDE.GetFieldByName("tgl_mulai") = Now
       MyDDE.GetFieldByName("tgl_selesai") = Now
       MyDDE.GetFieldByName("ph") = "7"
       txt(1).Text = frmProduksi.txtBox(5).Text
       txt(3).Text = ""
       tgl(0).Value = Now
       tgl(1).Value = Now
       tgl(2).Value = Now
    
       
    Case tmbDetail:
       MyDDE.ChildRecordset.Fields("tanggal_mulai") = Now
       MyDDE.ChildRecordset.Fields("tanggal_selesai") = Now
       MyDDE.ChildRecordset.Fields("QtyKering") = 0
       MyDDE.ChildRecordset.Fields("Qtybasah") = 0
       MyDDE.ChildRecordset.Fields("rendemen") = 0
       MyDDE.ChildRecordset.Fields("kekeringan") = False
       MyDDE.ChildRecordset.Fields("dryer") = TreeView1.SelectedItem.Key
       
       MyDDE.ChildRecordset.MoveFirst
       x = 1
       While Not MyDDE.ChildRecordset.EOF
         MyDDE.ChildRecordset.Fields("no") = x
         x = x + 1
         MyDDE.ChildRecordset.MoveNext
       Wend
       MyDDE.ChildRecordset.MoveLast
       
    Case tmbSave:
        If MyDDE.IsChildMemberReady = True Then
            MyDDE.ChildRecordset.Filter = adFilterNone
            Set GridDry.DataSource = MyDDE.ChildRecordset
            SaveToMO
            simpan_detail
            If Not MEdit Then
               SendDataToServer "update StatusProduksi set status=1,posisi='DRYER' where NoEkstraksi='" & txt(3).Text & "'"
            End If
            MEdit = False
            TreeView1_NodeClick TreeView1.SelectedItem
        End If
        
    Case tmbPrint:
      Dim lPrint As New utility
      lPrint.CallReportView "select * from v_dryer where noEkstraksi='" & txt(3).Text & "'", "jemur_dryer.rpt", ReportPath, "Dryer & Penjemuran"
      Set lPrint = Nothing

End Select
Exit Sub
1:
MessageBox Err.Description, "frmdryer:mydde_afterprepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 2
   Select Case AdReasonActiveDb
      Case tmbDelete:
         If Not MyDDE.ChildRecordset.Recordcount > 0 Then
            SendDataToServer "Update StatusProduksi set posisi='JEMUR', status=1 where noEkstraksi='" & txt(3).Text & "'"
         End If
   End Select
Exit Sub
2:
MessageBox Err.Description, "frmdryer:mydde_executiveorder" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  loadDetail
  TreeView1_NodeClick TreeView1.SelectedItem
End Sub

Function PrepareSQL()
On Error GoTo xErr
   MyDDE.PrepareAppend = "insert into dryer_header(id,orderID,noEkstraksi,grup,tanggal_mulai,tanggal_selesai,[keterangan]," & _
                                                   "tanggal,issued_by) " & _
                         "values ('" & txt(0).Text & "', '" & _
                                       txt(1).Text & "', '" & _
                                       txt(3).Text & "', '" & _
                                       txt(2).Text & "', '" & _
                                       Format(tgl(0).Value, "yyyy-MM-dd hh:mm:ss") & "', '" & _
                                       Format(tgl(1).Value, "yyyy-MM-dd hh:mm:ss") & "', '" & _
                                       txt(6).Text & "', '" & _
                                       Format(tgl(2).Value, "yyyy-MM-dd") & "', '" & _
                                       MainMenu.StatusBar1.Panels(1).Text & "')"
                       
   MyDDE.PrepareUpdate = "Update dryer_header set noEkstraksi = '" & txt(3).Text & "'," & _
                                                   "OrderID ='" & txt(1).Text & "'," & _
                                                   "grup='" & txt(2).Text & "'," & _
                                                   "tanggal ='" & Format(tgl(2).Value, "yyyy-MM-dd") & "'," & _
                                                   "tanggal_mulai='" & Format(tgl(0).Value, "yyyy-MM-dd hh:mm:ss") & "'," & _
                                                   "tanggal_selesai='" & Format(tgl(1).Value, "yyyy-MM-dd hh:mm:ss") & "'," & _
                                                   "keterangan ='" & MyDDE.GetFieldByName("keterangan") & "'," & _
                                                   "issued_by ='" & MainMenu.StatusBar1.Panels(1).Text & "' " & _
                         " where ID ='" & MyDDE.GetFieldByName("ID") & "'"
    
  MyDDE.PrepareDelete = "delete from dryer_header where id = '" & MyDDE.GetFieldByName("id") & "' "
Exit Function
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Function


Function simpan_detail()
On Error GoTo xErr
   With MyDDE.ChildRecordset
      If .Recordcount <> 0 Then
          .MoveFirst
            If SendDataToServer(" delete from [dryer_DETAIL] where (id = '" & MyDDE.GetFieldByName("id") & "')") = True Then
            Do
              If .EOF = True Then Exit Do
              SendDataToServer "insert into dryer_DETAIL (id,no,tanggal_mulai,tanggal_selesai,qtyKering,Dryer,qtyBasah,rendemen,kekeringan) " & _
                               " values ('" & txt(0).Text & "','" & _
                                          .Fields("no") & "','" & _
                                          Format(.Fields("tanggal_mulai"), "yyyy-MM-dd hh:mm:ss") & "','" & _
                                          Format(.Fields("tanggal_selesai"), "yyyy-MM-dd hh:mm:ss") & "'," & _
                                          FQty(.Fields("qtykering")) & ",'" & _
                                          .Fields("dryer") & "'," & _
                                          FQty(.Fields("QtyBasah")) & "," & _
                                          FQty(.Fields("rendemen")) & "," & _
                                          IIf(.Fields("kekeringan") = True, "1", "0") & ")"
             .MoveNext
           Loop
           End If
           .MoveLast
           GridDry.Refresh
           End If
   End With
Exit Function
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Function


Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 3
   PrepareSQL
   cmdLink(0).Enabled = False
   cmdLink(1).Enabled = False
   Select Case AdReasonActiveDb
      Case tmbAddNew:
         cmdLink(0).Enabled = True
         cmdLink(1).Enabled = True
         MEdit = False
      Case tmbEdit:
         cmdLink(0).Enabled = True
         cmdLink(1).Enabled = True
         MEdit = True
      Case tmbSave:
         MyDDE.IsChildMemberReady = True
      Case tmbDelete:
         Dim rsCek As New DBQuick
         rsCek.DBOpen "select posisi from StatusProduksi where NoEkstraksi='" & txt(3).Text & "'", CNN
         If rsCek.DBRecordset.Recordcount > 0 Then
            If rsCek.DBRecordset.Fields(0) <> "DRYER" Then
               MessageBox "Data Tidak bisa dihapus, sedang diproses ditempat lain", "Stop", msgOkOnly, msgCrtical
               MyDDE.CancelTrans = True
            End If
         Else
            MessageBox "Data Tidak Bisa Dihapus", "Error", msgOkOnly, msgCrtical
            MyDDE.CancelTrans = True
         End If
         rsCek.CloseDB
   End Select
Exit Sub
3:
MessageBox Err.Description, "frmdryer:mydde_afterprepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Form_Load()

   If lMode = 0 Then
      MyDDE.SetReadOnlyMode = False
   Else
      MyDDE.SetReadOnlyMode = True
   End If
   
   With MyDDE
      LoadDryer
      Set .BindForm = Me
          .BindFormTAG = "dryer"
      Set .ActiveConnection = CNN
          .PrepareQuery = "SELECT dryer_header.ID, dryer_header.NoEkstraksi, dryer_header.OrderID, dryer_header.grup, " & _
                           "dryer_header.tanggal_mulai,dryer_header.Keterangan , dryer_header.Tanggal, " & _
                           "dryer_header.issued_by FROM dryer_header INNER JOIN " & _
                          "[Manufacture Order] ON dryer_header.OrderID = [Manufacture Order].OrderID " & _
                          " WHERE [Manufacture Order].Status " & IIf(lMode = 0, "=", "<>") & " 'RELEASED'"
   End With
   HiasFormManTell Picture2, Me
   GridDry.RowHeight = 315
   GridDry.HeadLines = 2
   Set mCall = New frmCaller
End Sub


Private Sub tgl_Change(Index As Integer)
   If tgl(0).Value > tgl(1).Value Then
      MessageBox "Waktu Mulai tidak boleh lebis besar dar waktu selesai", "Peringatan", msgOkOnly, msgCrtical
   End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
'   On Error GoTo Hell
         MyDDE.ChildRecordset.Filter = "dryer = '" & Node.Key & "'"
      lblDryer.Caption = Node.Text
'   Exit Sub
'Hell:
'   If Err.Number = 3001 Then Err.Clear Else MessageBox Err.Description
End Sub

Private Sub txt_LostFocus(Index As Integer)
   If Index = 3 Then
      Dim rsCek As New DBQuick
      rsCek.DBOpen "select * from statusProduksi where noEkstraksi='" & txt(3).Text & "'", CNN, lckLockBatch
      If rsCek.DBRecordset.Recordcount > 0 Then
         rsCek.DBOpen "select * from dryer_header where noEkstraksi='" & txt(3).Text & "'", CNN, lckLockBatch
         If rsCek.DBRecordset.Recordcount > 0 Then
            MessageBox "Nomor Ekstraksi Ini Sudah Diinput...!", "Peringatan", msgOkOnly, msgCrtical
            txt(3).Text = ""
         End If
      Else
         MessageBox "Nomor Ekstraksi Ini tidak ditemukan...!", "Peringatan", msgOkOnly, msgCrtical
         txt(3).Text = ""
      End If
      rsCek.CloseDB
   End If

End Sub

Private Sub viewDate_Change()
   Select Case currCol
      Case 2: MyDDE.ChildRecordset.Fields("tanggal_mulai") = viewDate.Value
      Case 3: MyDDE.ChildRecordset.Fields("tanggal_selesai") = viewDate.Value
   End Select
End Sub


Private Sub SaveToMO()
   Dim dStart As Date
   Dim dFinish As Date
   Dim ActualTime As Double
   Dim sWCID As String
   Dim rsCek As New DBQuick
On Error GoTo xErr
   dStart = tgl(0).Value
   dFinish = tgl(1).Value
   ActualTime = Val(SelisihHariJam(dStart, dFinish, 2))
   
   rsCek.DBOpen "select WCID from WCenter_Header where FormID = 34", CNN
   If rsCek.DBRecordset.Recordcount > 0 Then
      sWCID = rsCek.DBRecordset.Fields(0)
      SendDataToServer "update [order output detail] set actual_time=" & ActualTime & " where OrderID='" & MyDDE.GetFieldByName("OrderID") & "' and WCID='" & sWCID & "'"
   End If
   rsCek.CloseDB
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

