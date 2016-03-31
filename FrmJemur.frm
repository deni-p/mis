VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmJemur 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penjemuran"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   9900
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmJemur.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   9900
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5355
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   1005
      BindFormTAG     =   "jemur"
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
      Height          =   5325
      Left            =   0
      ScaleHeight     =   5325
      ScaleWidth      =   9900
      TabIndex        =   16
      Top             =   0
      Width           =   9900
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAAF6F&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   2235
         TabIndex        =   31
         Top             =   3030
         Width           =   1650
         Begin VB.OptionButton Option2 
            BackColor       =   &H00EAAF6F&
            Caption         =   "Kering"
            Height          =   225
            Left            =   30
            TabIndex        =   33
            Top             =   45
            Width           =   780
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00EAAF6F&
            Caption         =   "Basah"
            Height          =   225
            Left            =   840
            TabIndex        =   32
            Top             =   45
            Width           =   765
         End
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "rendemen"
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
         Index           =   7
         Left            =   2235
         TabIndex        =   28
         Tag             =   "jemur"
         Top             =   2655
         Width           =   1425
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "OrderID"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   5
         Left            =   4305
         Locked          =   -1  'True
         TabIndex        =   2
         Tag             =   "jemur"
         Top             =   165
         Width           =   1950
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   6255
         Picture         =   "FrmJemur.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "jemur"
         Top             =   540
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   6255
         Picture         =   "FrmJemur.frx":6BDC
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "jemur"
         Top             =   180
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame Opsi 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   7950
         TabIndex        =   26
         Top             =   4155
         Visible         =   0   'False
         Width           =   1590
         Begin VB.OptionButton Opt2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Kotor"
            Height          =   225
            Left            =   825
            TabIndex        =   15
            Top             =   30
            Width           =   765
         End
         Begin VB.OptionButton Opt1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Bersih"
            Height          =   225
            Left            =   45
            TabIndex        =   14
            Top             =   30
            Width           =   780
         End
      End
      Begin MSDataGridLib.DataGrid GridJemur 
         Height          =   3645
         Left            =   4680
         TabIndex        =   13
         Top             =   1170
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   6429
         _Version        =   393216
         HeadLines       =   1
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "lajur"
            Caption         =   "Lajur"
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
            DataField       =   "kebersihan"
            Caption         =   "Kebersihan Screen"
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
            DataField       =   "kondisi"
            Caption         =   "Kondisi Screen"
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
               Alignment       =   2
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Locked          =   -1  'True
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "QtyKering"
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
         Index           =   4
         Left            =   2235
         TabIndex        =   11
         Tag             =   "jemur"
         Top             =   2310
         Width           =   1425
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "keterangan"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   885
         Index           =   6
         Left            =   225
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Tag             =   "jemur"
         Top             =   3930
         Width           =   4050
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "qtyBasah"
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
         Index           =   3
         Left            =   2235
         TabIndex        =   10
         Tag             =   "jemur"
         Top             =   1950
         Width           =   1425
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
         Left            =   8160
         TabIndex        =   7
         Tag             =   "jemur"
         Top             =   525
         Width           =   1470
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "no_ekstraksi"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   1
         Left            =   4305
         TabIndex        =   4
         Tag             =   "jemur"
         Top             =   525
         Width           =   1950
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
         Left            =   765
         Locked          =   -1  'True
         TabIndex        =   1
         Tag             =   "jemur"
         Top             =   165
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "tgl_mulai"
         DataSource      =   "dde"
         Height          =   315
         Index           =   0
         Left            =   2205
         TabIndex        =   8
         Tag             =   "jemur"
         Top             =   1215
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy   hh:mm"
         Format          =   16580611
         CurrentDate     =   39419
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "tanggal"
         DataSource      =   "dde"
         Height          =   315
         Index           =   2
         Left            =   8160
         TabIndex        =   6
         Tag             =   "jemur"
         Top             =   165
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   16580611
         CurrentDate     =   39365
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "tgl_selesai"
         DataSource      =   "dde"
         Height          =   315
         Index           =   1
         Left            =   2220
         TabIndex        =   9
         Tag             =   "jemur"
         Top             =   1575
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy   hh:mm"
         Format          =   16580611
         CurrentDate     =   39419
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7590
         TabIndex        =   34
         Top             =   4890
         Width           =   2025
      End
      Begin VB.Line Line1 
         Index           =   11
         X1              =   8775
         X2              =   5850
         Y1              =   5205
         Y2              =   5205
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Hasil Kekeringan                                                     "
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   30
         Top             =   3060
         Width           =   4065
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   3120
         X2              =   195
         Y1              =   3315
         Y2              =   3315
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Hasil Rendemen                                                     %"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   29
         Top             =   2715
         Width           =   4065
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   3120
         X2              =   195
         Y1              =   2970
         Y2              =   2970
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   4755
         X2              =   2715
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   3120
         X2              =   195
         Y1              =   2625
         Y2              =   2625
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty Kering Lembaran Agar                                    Kg"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   25
         Top             =   2370
         Width           =   4065
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan :"
         Height          =   255
         Index           =   33
         Left            =   240
         TabIndex        =   23
         Top             =   3690
         Width           =   1050
      End
      Begin VB.Line Line1 
         Index           =   9
         X1              =   2235
         X2              =   195
         Y1              =   2265
         Y2              =   2265
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   9075
         X2              =   7035
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   9075
         X2              =   7035
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal && waktu selesai"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   21
         Top             =   1620
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal && waktu mulai"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   20
         Top             =   1260
         Width           =   1890
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   2040
         X2              =   330
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   255
         Index           =   12
         Left            =   375
         TabIndex        =   19
         Top             =   195
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         Height          =   255
         Index           =   2
         Left            =   7050
         TabIndex        =   17
         Top             =   585
         Width           =   2055
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   4755
         X2              =   2715
         Y1              =   840
         Y2              =   840
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
         Left            =   7050
         TabIndex        =   24
         Top             =   210
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty Basah Lembaran Agar                                    Kg"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   22
         Top             =   2010
         Width           =   3810
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Manufacture Order"
         Height          =   255
         Index           =   7
         Left            =   2745
         TabIndex        =   27
         Top             =   210
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "No Ekstraksi"
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   18
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By                                                  %"
         Height          =   255
         Index           =   10
         Left            =   5895
         TabIndex        =   35
         Top             =   4950
         Width           =   4065
      End
   End
End
Attribute VB_Name = "FrmJemur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RsDetail As New DBQuick
Private RsLookup As New DBQuick
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private MEdit As Boolean
Private lMode As Byte

Public Property Let SetMode(ByVal Value As Byte)
   lMode = Value
End Property


Private Sub loadDetail()
    RsDetail.DBOpen "select lajur,kebersihan,kondisi from jemur_detail where id = '" & MyDDE.GetFieldByName("id") & "'", CNN
    Set MyDDE.ChildRecordset = RsDetail.DBRecordset
    Set GridJemur.DataSource = MyDDE.ChildRecordset
End Sub



Private Sub cmdLink_Click(Index As Integer)
   Select Case Index
      Case 0: RsLookup.DBOpen "select OrderID,OrderNAme,Type, requireDate from [manufacture Order] where status='RELEASED'", CNN
      Case 1: RsLookup.DBOpen "select noEkstraksi from statusProduksi where posisi='CUTTER' and status=1", CNN
   End Select
   
   If RsLookup.Recordcount > 0 Then
      Set mCall.FormData = RsLookup.DBRecordset
      If Index = 0 Then
         mCall.FromTagActive = "Manufacture Order"
      Else
         mCall.FromTagActive = "No Ekstraksi"
      End If
   Else
      MessageBox "Data Tidak Tersedia", "Peringatan", msgOkOnly, msgCrtical
   End If
   
End Sub

Private Sub GridJemur_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo Hell
   Opsi.Visible = False
   Select Case GridJemur.col
      Case 1: Opsi.Visible = True
               Opt1.Caption = "Bersih"
               Opt2.Caption = "Kotor"
               Opsi.Move GridJemur.Left + GridJemur.Columns(1).Left, _
                         GridJemur.Top + GridJemur.RowTop(GridJemur.row), _
                         GridJemur.Columns(1).width, _
                         GridJemur.RowHeight
      Case 2: Opsi.Visible = True
               Opt1.Caption = "Baik"
               Opt2.Caption = "Rusak"
               Opsi.Move GridJemur.Left + GridJemur.Columns(2).Left, _
                         GridJemur.Top + GridJemur.RowTop(GridJemur.row), _
                         GridJemur.Columns(2).width, _
                         GridJemur.RowHeight
   End Select
Exit Sub
Hell:
   If Err.Number = 6148 Then
      Err.Clear
   Else
      MessageBox Err.Description, "Error", msgOkOnly, msgExclamation
   End If
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
   Select Case UCase(mCall.FromTagActive)
      Case "MANUFACTURE ORDER": txt(5).Text = mCall.GetFieldByName(0)
      Case "NO EKSTRAKSI": txt(1).Text = mCall.GetFieldByName(0)
   End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim x As Integer
On Error GoTo xErr
Select Case AdReasonActiveDb
    Case tmbAddNew
       txt(0).Text = IndexAuto
       MyDDE.GetFieldByName("keterangan") = "-"
       MyDDE.GetFieldByName("tanggal") = Now
       MyDDE.GetFieldByName("tgl_mulai") = Now
       MyDDE.GetFieldByName("tgl_selesai") = Now
       tgl(0).Value = Now
       tgl(1).Value = Now
       tgl(2).Value = Now
       txt(5).Text = frmProduksi.txtBox(5).Text
       'txt(1).Text = frmProduksi.txtBox(1).Text
       txt(3).Text = 0
       txt(4).Text = 0
       
    Case tmbDetail:
       MyDDE.ChildRecordset.Fields("kebersihan") = "Bersih"
       MyDDE.ChildRecordset.Fields("kondisi") = "Baik"
       
    Case tmbSave:
        If MyDDE.IsChildMemberReady = True Then
            SendDataToServer "update StatusProduksi set status=1,posisi='JEMUR' where NoEkstraksi='" & MyDDE.GetFieldByName("no_ekstraksi") & "'"
            SaveToMO
            simpan_detail
        End If
    
    Case tmbPrint:
      Dim lPrint As New utility
      lPrint.CallReportView "select * from v_dryer where noEkstraksi='" & txt(1).Text & "'", "jemur_dryer.rpt", ReportPath, "Dryer & Penjemuran"
      Set lPrint = Nothing
End Select
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  loadDetail
  If MyDDE.GetFieldByName("kekeringan") = True Then
      Option2.Value = True
      Option1.Value = False
  Else
      Option2.Value = False
      Option1.Value = True
  End If
  Label1.Caption = IIf(IsNull(MyDDE.GetFieldByName("approved_by")), "", MyDDE.GetFieldByName("approved_by"))
End Sub

Function PrepareSQL()
On Error GoTo xErr
   MyDDE.PrepareAppend = "insert into jemur_header(id,no_ekstraksi,OrderID,tanggal,grup,tgl_mulai,tgl_selesai,QtyBasah,rendemen,kekeringan," & _
                                                 " QtyKering,[keterangan],issued_by) " & _
                         "values ('" & txt(0).Text & "', '" & _
                                 txt(1).Text & "', '" & _
                                 txt(5).Text & "','" & _
                                 Format(tgl(2).Value, "yyyy-MM-dd") & "','" & _
                                 MyDDE.GetFieldByName("grup") & "','" & _
                                 Format(tgl(0).Value, "yyyy-MM-dd hh:mm:ss") & "', '" & _
                                 Format(tgl(1).Value, "yyyy-MM-dd hh:mm:ss") & "', " & _
                                 FQty(MyDDE.GetFieldByName("qtyBasah")) & "," & _
                                 FQty(MyDDE.GetFieldByName("rendemen")) & "," & _
                                 IIf(Option2.Value = True, "1", "0") & "," & _
                                 FQty(txt(4).Text) & ",'" & _
                                 MyDDE.GetFieldByName("keterangan") & "','" & _
                                 MainMenu.StatusBar1.Panels(1).Text & "')"
                       
   MyDDE.PrepareUpdate = "Update jemur_header set no_ekstraksi = '" & MyDDE.GetFieldByName("no_ekstraksi") & "'," & _
                                                   "OrderID='" & MyDDE.GetFieldByName("OrderID") & "'," & _
                                                   "tanggal ='" & Format(tgl(2).Value, "yyyy-MM-dd") & "'," & _
                                                   "grup ='" & MyDDE.GetFieldByName("grup") & "'," & _
                                                   "tgl_mulai='" & Format(tgl(0).Value, "yyyy-MM-dd hh:mm:ss") & "'," & _
                                                   "tgl_selesai='" & Format(tgl(1).Value, "yyyy-MM-dd hh:mm:ss") & "'," & _
                                                   "QtyBasah =" & MyDDE.GetFieldByName("QtyBasah") & "," & _
                                                   "rendemen =" & MyDDE.GetFieldByName("rendemen") & "," & _
                                                   "kekeringan =" & IIf(Option2.Value = True, "1", "0") & "," & _
                                                   "QtyKering =" & FQty(txt(4).Text) & "," & _
                                                   "keterangan ='" & MyDDE.GetFieldByName("keterangan") & "'," & _
                                                   "issued_by ='" & MainMenu.StatusBar1.Panels(1).Text & "' " & _
                           "where ID ='" & MyDDE.GetFieldByName("ID") & "'"

   MyDDE.PrepareDelete = "delete from jemur_header where id'" & MyDDE.GetFieldByName("ID") & "'"

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
         If SendDataToServer(" delete from [JEMUR_DETAIL] where (id = '" & MyDDE.GetFieldByName("id") & "')") = True Then
         Do
           If .EOF = True Then Exit Do
           SendDataToServer "insert into JEMUR_DETAIL (id,lajur,kebersihan,kondisi) " & _
           " values ('" & txt(0).Text & "','" & .Fields("lajur") & "','" & _
           .Fields("kebersihan") & "','" & .Fields("kondisi") & "')"
          .MoveNext
        Loop
        End If
        .MoveLast
        GridJemur.Refresh
        End If
    End With
    
Exit Function
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Function

Private Function IndexAuto() As String
Dim Rc As New DBQuick
Dim TglSaiki As String
Dim Inom As Long
TglSaiki = Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2)
Rc.DBOpen "SELECT MAX(RIGHT(id, 5)) AS MaxNom FROM [jemur_header] WHERE (GETDATE() = { fn NOW() })", CNN, lckLockReadOnly
With Rc
     If .DBRecordset.Recordcount <> 0 Then
        Inom = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + 1
     Else
        Inom = 1
     End If
     Select Case Len(Trim(Str(Inom)))
            Case 0: IndexAuto = "PD/" & TglSaiki & "-" & Trim(Str(Inom))
            Case 1: IndexAuto = "PD/" & TglSaiki & "-" & "0000" & Trim(Str(Inom))
            Case 2: IndexAuto = "PD/" & TglSaiki & "-" & "000" & Trim(Str(Inom))
            Case 3: IndexAuto = "PD/" & TglSaiki & "-" & "00" & Trim(Str(Inom))
            Case 4: IndexAuto = "PD/" & TglSaiki & "-" & "0" & Trim(Str(Inom))
     End Select
End With
End Function


Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
   cmdLink(0).Enabled = False
   cmdLink(1).Enabled = False
   PrepareSQL
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
         
   End Select
End Sub


Private Sub Form_Load()
   If lMode = 0 Then
      MyDDE.SetReadOnlyMode = False
   Else
      MyDDE.SetReadOnlyMode = True
   End If


With MyDDE
   Set .BindForm = Me
       .BindFormTAG = "jemur"
   Set .ActiveConnection = CNN
   
       .PrepareQuery = " SELECT JEMUR_HEADER.id, JEMUR_HEADER.no_ekstraksi, JEMUR_HEADER.orderID, JEMUR_HEADER.tanggal, " & _
                        "JEMUR_HEADER.grup, JEMUR_HEADER.tgl_mulai, JEMUR_HEADER.tgl_selesai, JEMUR_HEADER.qtyBasah,JEMUR_HEADER.rendemen,JEMUR_HEADER.kekeringan, " & _
                        "JEMUR_HEADER.QtyKering,JEMUR_HEADER.Keterangan , JEMUR_HEADER.issued_by " & _
                       " FROM [Manufacture Order] INNER JOIN " & _
                       " JEMUR_HEADER ON dbo.[Manufacture Order].OrderID = dbo.JEMUR_HEADER.orderID " & _
                       " WHERE (dbo.[Manufacture Order].Status " & IIf(lMode = 0, " = ", " <> ") & "'RELEASED') "
   End With
   HiasFormManTell Picture2, Me
   GridJemur.RowHeight = 300
   Set mCall = New frmCaller
End Sub

Private Sub Opt1_Click()
   If Opt1.Value = True Then
      If GridJemur.col = 1 Then
         MyDDE.ChildRecordset.Fields("kebersihan") = Opt1.Caption
      ElseIf GridJemur.col = 2 Then
         MyDDE.ChildRecordset.Fields("kondisi") = Opt1.Caption
      End If
   End If
End Sub

Private Sub Opt2_Click()
   If Opt2.Value = True Then
      If GridJemur.col = 1 Then
         MyDDE.ChildRecordset.Fields("kebersihan") = Opt2.Caption
      ElseIf GridJemur.col = 2 Then
         MyDDE.ChildRecordset.Fields("kondisi") = Opt2.Caption
      End If
   End If
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
   
   rsCek.DBOpen "select WCID from WCenter_Header where FormID = 47", CNN
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

Private Sub tgl_Change(Index As Integer)
   If tgl(0).Value > tgl(1).Value Then
      MessageBox "Waktu Mulai tidak boleh lebih besar dari waktu selesai"
   End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
   On Error GoTo xErr
   If Index = 4 Or Index = 3 Then
      If Val(txt(3).Text) < Val(txt(4).Text) Then
         MessageBox " Qty Lembaran agar Kering harus lebih kecil dari Qty Lembaran Agar Basah", "Peringatan", msgOkOnly, msgCrtical
      End If
      txt(7).Text = Val(txt(4)) / Val(txt(3)) * 100
   End If
   
   If Index = 1 Then
      Dim rsCek As New DBQuick
      rsCek.DBOpen "select * from statusProduksi where noEkstraksi='" & txt(1).Text & "'", CNN, lckLockBatch
      If rsCek.DBRecordset.Recordcount > 0 Then
         rsCek.DBOpen "select * from JEMUR_HEADER where no_Ekstraksi='" & txt(1).Text & "'", CNN, lckLockBatch
         If rsCek.DBRecordset.Recordcount > 0 Then
            MessageBox "Nomor Ekstraksi Ini Sudah Diinput...!", "Peringatan", msgOkOnly, msgCrtical
            txt(1).Text = ""
         End If
      Else
         MessageBox "Nomor Ekstraksi Ini tidak ditemukan...!", "Peringatan", msgOkOnly, msgCrtical
         txt(1).Text = ""
      End If
      rsCek.CloseDB
   End If
   
Exit Sub
xErr:
   MessageBox "Data Tidak Sesuai", "Error", msgOkOnly, msgExclamation
   Err.Clear
End Sub
