VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmMPermintaanSample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Permintaan Sample"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPermintaanSample.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   11400
   Begin SemeruDC.SemeruOleDC MYDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5520
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   1005
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
   Begin VB.PictureBox Picture2 
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
      Height          =   5520
      Left            =   0
      ScaleHeight     =   5520
      ScaleWidth      =   11505
      TabIndex        =   1
      Top             =   0
      Width           =   11505
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "approved_by"
         Height          =   330
         Index           =   3
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   12
         Tag             =   "minta_sample"
         Top             =   5160
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Nomor"
         Height          =   330
         Index           =   0
         Left            =   1005
         TabIndex        =   7
         Tag             =   "minta_sample"
         Top             =   120
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "FullName"
         Height          =   330
         Index           =   1
         Left            =   1005
         TabIndex        =   5
         Tag             =   "minta_sample"
         Top             =   855
         Width           =   2895
      End
      Begin VB.CommandButton cmdLink 
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
         Height          =   315
         Left            =   3895
         MaskColor       =   &H000000C0&
         Picture         =   "frmPermintaanSample.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   863
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "empID"
         Height          =   330
         Index           =   2
         Left            =   1005
         TabIndex        =   2
         Tag             =   "minta_sample"
         Top             =   1215
         Visible         =   0   'False
         Width           =   2895
      End
      Begin MSComCtl2.MonthView dt2 
         Height          =   2370
         Left            =   2025
         TabIndex        =   3
         Top             =   1995
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   57802753
         CurrentDate     =   39395
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3255
         Left            =   195
         TabIndex        =   6
         Tag             =   "minta_sampel"
         Top             =   1740
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   5741
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "ItemName"
            Caption         =   "Nama Produk"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "M/d/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Jumlah"
            Caption         =   "Jumlah"
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
            DataField       =   "tanggal_butuh"
            Caption         =   "Tanggal dibutuhkan"
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
         BeginProperty Column03 
            DataField       =   "companyname"
            Caption         =   "Ditujukan ke"
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
            DataField       =   "keterangan"
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
         BeginProperty Column05 
            DataField       =   "noitem"
            Caption         =   "noitem"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "partnerID"
            Caption         =   "partnerID"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   2520
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Button          =   -1  'True
            EndProperty
            BeginProperty Column03 
               Button          =   -1  'True
               ColumnWidth     =   2160
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2520
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "tanggal"
         Height          =   330
         Index           =   0
         Left            =   1005
         TabIndex        =   8
         Tag             =   "minta_sample"
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   57802755
         CurrentDate     =   39394
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   5235
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales"
         Height          =   195
         Index           =   2
         Left            =   285
         TabIndex        =   11
         Top             =   930
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   195
         Index           =   1
         Left            =   285
         TabIndex        =   10
         Top             =   555
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor"
         Height          =   195
         Index           =   0
         Left            =   285
         TabIndex        =   9
         Top             =   195
         Width           =   465
      End
      Begin VB.Line Line1 
         X1              =   1300
         X2              =   255
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Line Line2 
         X1              =   1815
         X2              =   255
         Y1              =   780
         Y2              =   795
      End
      Begin VB.Line Line3 
         X1              =   1875
         X2              =   255
         Y1              =   1170
         Y2              =   1170
      End
   End
End
Attribute VB_Name = "frmMPermintaanSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsDetail As DBQuick
Private WithEvents mCall  As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private WithEvents mPRO As frmCaller
Attribute mPRO.VB_VarHelpID = -1
Private WithEvents mCos As frmCaller
Attribute mCos.VB_VarHelpID = -1
Dim IDGen As New IDGenerator
Dim rsSAles As New DBQuick
Dim rsproduk As New DBQuick
Dim rsCostumer As New DBQuick

Private Sub Command1_Click()

   
End Sub

Private Sub cmdLink_Click()
'   rsSAles.DBOpen "select * from employees", CNN
'
'   Set mCall = New frmCaller
'   Set mCall.FormData = rsSAles.DBRecordset
'   mCall.FromTagActive = "Sales"
'   mCall.CaptionLink = "Sales"
    Tsample = True
    frmCallerBaru.Show 1
End Sub

Private Sub DataGrid1_ButtonClick(ByVal ColIndex As Integer)
Select Case ColIndex
    Case 3
          '  rsCostumer.DBOpen "select * from Customer", CNN
            rsCostumer.DBOpen "select * from partnerdb", CNN
            Set mCos = New frmCaller
            Set mCos.FormData = rsCostumer.DBRecordset
            mCos.FromTagActive = "Costumer"
            mCos.CaptionLink = "Costumer"
            dt2.Visible = False
    Case 2
        dt2.Visible = True
        dt2.Move DataGrid1.Columns(2).Left + 100, (DataGrid1.RowTop(DataGrid1.row) + DataGrid1.Top)
 
End Select
 
End Sub
Private Sub dt2_DateClick(ByVal DateClicked As Date)
MyDDE.ChildRecordset.Fields("tanggal_butuh").Value = dt2.Value
dt2.Visible = False
End Sub

Private Sub Form_Load()
On Error Resume Next
HiasFormManTell Picture2, Me
DTPicker1(0).Value = Now
With MyDDE
    .EditModeReplace = False
     Set .BindForm = frmMPermintaanSample
    .BindFormTAG = "minta_sample"
     Set .ActiveConnection = CNN
    '.PrepareQuery = "select * from PermintaanSample, employees where PermintaanSample.EmpID = employees.EmpID"
     .PrepareQuery = "select nomor,tanggal,empid,fullname,approved_by from PermintaanSample" ', employees where PermintaanSample.EmpID = employees.EmpID"
End With

Set DTPicker1(0).DataSource = MyDDE.ActiveRecordset
DTPicker1(0).Enabled = False
DataGrid1.Columns(5).Visible = False
DataGrid1.Columns(6).Visible = False
End Sub

Private Sub header()
With MyDDE
  ' .PrepareAppend = " insert into PermintaanSample (nomor,tanggal,EmpID) values ('" & .GetFieldByName("nomor") & "', '" & .GetFieldByName("tanggal") & "', '" & .GetFieldByName("EmpID") & "')"
   
  
    .PrepareAppend = " INSERT INTO  permintaansample (nomor, tanggal,  empID, fullname, ordered_by) " & _
                     " VALUES (N'" & Text1(0).Text & "', convert(Datetime, '" & Format(DTPicker1(0).Value, "dd/mm/yy") & "',3) , N'" & Text1(2).Text & "', N'" & Text1(1).Text & "','" & MainMenu.StatusBar1.Panels(1).Text & "')"

    .PrepareUpdate = " UPDATE permintaansample " & _
                     " Set empID='" & Text1(2).Text & "', Tanggal = convert(Datetime, '" & Format(DTPicker1(0).Value, "dd/mm/yy") & "',3),fullname=N'" & Text1(1).Text & "', ordered_by='" & MainMenu.StatusBar1.Panels(1).Text & "'" & _
                     " WHERE (nomor = N'" & Text1(0).Text & "')"

   .PrepareDelete = " DELETE FROM permintaansample WHERE (nomor = N'" & Text1(0).Text & "')"
End With
Err.Clear
End Sub

Private Sub detail()
With MyDDE.ChildRecordset
   If .Recordcount <> 0 Then
       .MoveFirst
        If SendDataToServer(" delete from permintaan_sample_detail where (nomor = '" & MyDDE.GetFieldByName("nomor") & "')") = True Then
         Do
           If .EOF = True Then Exit Do
           SendDataToServer "insert into permintaan_sample_detail (nomor,NoItem,Jumlah,Tanggal_butuh,partnerID, keterangan)" & _
           " values ('" & Text1(0).Text & "', " & _
           " '" & DataGrid1.Columns(5).Text & "', " & _
           " '" & DataGrid1.Columns(1).Text & "', " & _
           " convert(Datetime, '" & Format(DataGrid1.Columns(2).Value, "dd/mm/yy") & "',3), " & _
           " '" & DataGrid1.Columns(6).Text & "', " & _
           " '" & DataGrid1.Columns(4).Text & "')"
          .MoveNext
        Loop
        End If
        .MoveLast
        DataGrid1.Refresh
      End If
    End With
End Sub

Private Sub listTipeItem_Click()
'   MYDDE.ChildRecordset.Fields("Costumer") = listTipeItem.Text
'   listTipeItem.Visible = False
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
    MyDDE.GetFieldByName("FullName") = rsSAles.DBRecordset.Fields("FullName")
    MyDDE.GetFieldByName("EmpID") = rsSAles.DBRecordset.Fields("EmpID")
End Sub


Private Sub mCos_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
MyDDE.ChildRecordset.Fields("partnerid") = mCos.GetFieldByName(0)
MyDDE.ChildRecordset.Fields("companyname") = mCos.GetFieldByName(1)
End Sub

Private Sub mPRO_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Debug.Print
    MyDDE.ChildRecordset.Fields("noitem") = mPRO.GetFieldByName(0)
    MyDDE.ChildRecordset.Fields("itemname") = mPRO.GetFieldByName(1)
End Sub


Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
Select Case AdReasonActiveDb
       Case tmbAddNew:
             cmdLink.Enabled = True
             DTPicker1(0).Enabled = True
            MyDDE.GetFieldByName("nomor") = IDGen.GetID("SAMPLE")
       Case tmbEdit:
            cmdLink.Enabled = True
            Set DTPicker1(0).DataSource = MyDDE.ActiveRecordset
            DTPicker1(0).Enabled = True
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               detail
               DTPicker1(0).Enabled = False
               cmdLink.Enabled = False
            End If
       Case tmbCancel:
            cmdLink.Enabled = False
            DTPicker1(0).Enabled = False
        Case tmbDelete:
            cmdLink.Enabled = False
            DTPicker1(0).Enabled = False
       Case tmbDetail:
             rsproduk.DBOpen "select NoItem,ItemName,UOM from inventory WHERE (Manufacture = 1) ORDER BY NoItem", CNN
             Set mPRO = New frmCaller
             Set mPRO.FormData = rsproduk.DBRecordset
             mPRO.FromTagActive = "Produk"
             mPRO.CaptionLink = "Produk"
        
       Case tmbPrint:
            Dim aReport As New utility
            aReport.CallReportView "select * from permintaan_sample where nomor='" & Text1(0) & "'", "permintaan sampel.rpt", ReportPath, "Permintaan Sample"
            Set aReport = Nothing
       Case tmbQuit:
'            Unload Me
End Select
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
header
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   Set RsDetail = New DBQuick
  ' rsdetail.DBOpen "select * from minta_sample where Nomor = '" & MyDDE.GetFieldByName("nomor") & "'", CNN

   RsDetail.DBOpen "select permintaan_sample_detail.noitem, inventory.itemname,permintaan_sample_detail.jumlah,permintaan_sample_detail.tanggal_butuh,permintaan_sample_detail.keterangan,dbo.Permintaan_sample_detail.PartnerId , dbo.PartnerDB.CompanyName from Permintaan_sample_detail INNER JOIN Inventory ON Permintaan_sample_detail.NoItem = Inventory.NoItem inner join dbo.PartnerDB ON dbo.Permintaan_sample_detail.partnerID = dbo.PartnerDB.PartnerID where Permintaan_sample_detail.Nomor = '" & MyDDE.GetFieldByName("nomor") & "'", CNN
   Set MyDDE.ChildRecordset = RsDetail.DBRecordset
   Set DataGrid1.DataSource = MyDDE.ChildRecordset
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
Select Case AdReasonActiveDb
     Case tmbAddNew
         cmdLink.Enabled = True
         Set DTPicker1(0).DataSource = MyDDE.ActiveRecordset
'         MYDDE.GetFieldByName("tanggal") = DTPicker1(0).value
     Case tmbSave
          MyDDE.IsChildMemberReady = True
'         If MYDDE.IsChildMemberReady = True Then
'            detail
'         End If
     Case tmbDetail
     
'        rsproduk.DBOpen "select * from inventory", CNN
'        Set mPRO = New frmCaller
'        Set mPRO.FormData = rsproduk.DBRecordset
'        mPRO.FromTagActive = "Produk"
'        mPRO.CaptionLink = "Produk"

     
  '   Case tmbDelete
         
         
End Select

End Sub





