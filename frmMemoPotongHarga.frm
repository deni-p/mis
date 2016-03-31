VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmMemoPotongHarga 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memo Pemotongan Harga"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10155
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMemoPotongHarga.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   10155
   Begin SemeruDC.SemeruOleDC MYDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   3810
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   1005
      BindFormTAG     =   "memo"
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
      Height          =   3795
      Left            =   0
      ScaleHeight     =   3795
      ScaleWidth      =   10155
      TabIndex        =   17
      Top             =   0
      Width           =   10155
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "approved_by"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   11
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   32
         Tag             =   "memo"
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "No Faktur"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   1
         Left            =   1200
         TabIndex        =   29
         Tag             =   "memo"
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   8820
         MaskColor       =   &H000000C0&
         Picture         =   "frmMemoPotongHarga.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "SPPH"
         Top             =   1328
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   3255
         MaskColor       =   &H000000C0&
         Picture         =   "frmMemoPotongHarga.frx":6BDC
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "SPPH"
         Top             =   968
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   4935
         MaskColor       =   &H000000C0&
         Picture         =   "frmMemoPotongHarga.frx":6F66
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "SPPH"
         Top             =   608
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Revisi"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   9
         Left            =   6675
         TabIndex        =   4
         Tag             =   "memo"
         Top             =   1680
         Width           =   2475
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Doc No"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   8
         Left            =   6675
         TabIndex        =   13
         Tag             =   "memo"
         Top             =   1320
         Width           =   2145
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Marketing ID"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   7
         Left            =   6675
         TabIndex        =   12
         Tag             =   "memo"
         Top             =   960
         Width           =   2475
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Sales ID"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   6
         Left            =   6675
         TabIndex        =   11
         Tag             =   "memo"
         Top             =   600
         Width           =   2475
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "lampiran"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   5
         Left            =   6675
         TabIndex        =   10
         Tag             =   "memo"
         Top             =   240
         Width           =   2475
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "reason"
         DataSource      =   "DDE"
         Height          =   1260
         Index           =   4
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   9
         Tag             =   "memo"
         Top             =   2040
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Discount"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   3
         Left            =   1200
         TabIndex        =   8
         Tag             =   "memo"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "No Item"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   2
         Left            =   1200
         TabIndex        =   6
         Tag             =   "memo"
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "CompanyName"
         Height          =   330
         Index           =   0
         Left            =   1200
         TabIndex        =   2
         Tag             =   "memo"
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Memo ID"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   10
         Left            =   6675
         TabIndex        =   16
         Tag             =   "memo"
         Top             =   2370
         Width           =   2475
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   3255
         MaskColor       =   &H000000C0&
         Picture         =   "frmMemoPotongHarga.frx":72F0
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "SPPH"
         Top             =   1328
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "Date Memo"
         DataSource      =   "MYDDE"
         Height          =   300
         Index           =   0
         Left            =   1200
         TabIndex        =   1
         Tag             =   "memo"
         Top             =   255
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   57999363
         CurrentDate     =   39335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "Effective Date"
         DataSource      =   "MYDDE"
         Height          =   300
         Index           =   1
         Left            =   6675
         TabIndex        =   15
         Tag             =   "memo"
         Top             =   2025
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   57999363
         CurrentDate     =   39335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By"
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   33
         Top             =   3435
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Memo"
         Height          =   195
         Index           =   12
         Left            =   5550
         TabIndex        =   31
         Top             =   2438
         Width           =   420
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   11
         X1              =   6690
         X2              =   5505
         Y1              =   2310
         Y2              =   2310
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Faktur"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   1035
         Width           =   705
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   5
         X1              =   6690
         X2              =   5505
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   4
         X1              =   1290
         X2              =   105
         Y1              =   1995
         Y2              =   1995
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   1290
         X2              =   105
         Y1              =   1635
         Y2              =   1635
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   1290
         X2              =   105
         Y1              =   1275
         Y2              =   1275
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   1290
         X2              =   105
         Y1              =   915
         Y2              =   915
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   1290
         X2              =   105
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date"
         Height          =   195
         Index           =   11
         Left            =   5550
         TabIndex        =   28
         Top             =   2085
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Revisi"
         Height          =   195
         Index           =   10
         Left            =   5550
         TabIndex        =   27
         Top             =   1755
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Document"
         Height          =   195
         Index           =   9
         Left            =   5550
         TabIndex        =   26
         Top             =   1395
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marketing"
         Height          =   195
         Index           =   8
         Left            =   5550
         TabIndex        =   25
         Top             =   1035
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales"
         Height          =   195
         Index           =   7
         Left            =   5550
         TabIndex        =   24
         Top             =   675
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lampiran"
         Height          =   195
         Index           =   6
         Left            =   5550
         TabIndex        =   23
         Top             =   315
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alasan"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Top             =   2055
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   1755
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Item"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   1395
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   675
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   315
         Width           =   570
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   6
         X1              =   6690
         X2              =   5505
         Y1              =   915
         Y2              =   915
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   7
         X1              =   6690
         X2              =   5505
         Y1              =   1275
         Y2              =   1275
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   8
         X1              =   6690
         X2              =   5505
         Y1              =   1635
         Y2              =   1635
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   9
         X1              =   6690
         X2              =   5505
         Y1              =   1995
         Y2              =   1995
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   10
         X1              =   6690
         X2              =   5505
         Y1              =   2655
         Y2              =   2655
      End
   End
End
Attribute VB_Name = "frmMemoPotongHarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private WithEvents cos As frmCaller
Attribute cos.VB_VarHelpID = -1
Private WithEvents invo As frmCaller
Attribute invo.VB_VarHelpID = -1
Private WithEvents sale As frmCaller
Attribute sale.VB_VarHelpID = -1
Private WithEvents market As frmCaller
Attribute market.VB_VarHelpID = -1
Private RcPartner As New DBQuick
Dim sql As String
'Dim mcall As New frmCaller

Private Sub OpenPartner(Index As Integer, Optional Params As String)
'Dim mcall As New frmCaller
Set RcPartner = New DBQuick
Set mCall = New frmCaller
Select Case Index
       Case 0:
            RcPartner.DBOpen "SELECT PartnerID AS [Kode Cust], CompanyName as [Nama Perusahaan], Address AS [Alamat], City as Kota FROM  PartnerDb WHERE (PartnerType = 'CUSTOMER') ORDER BY CompanyName", CNN, lckLockReadOnly
            mCall.FromTagActive = "CUSTOMER"
            
       Case 1:
           ' RcPartner.DBOpen "SELECT TransID AS [No Faktur],  dateTrans AS Tanggal  FROM TransData WHERE TypeTrans ='AR' and noItem='" & Params & "' ORDER BY TransID", CNN, lckLockReadOnly
            RcPartner.DBOpen "SELECT TransID AS [No Faktur],  dateTrans AS Tanggal  FROM TransData WHERE TypeTrans ='AR' ORDER BY TransID", CNN, lckLockReadOnly
            mCall.FromTagActive = "FAKTUR"
       Case 2:
            '""
       Case 3:
       Case 4:
            RcPartner.DBOpen "SELECT [Detail TransData].NoItem as [Kode Barang],Inventory.ItemName as [Nama Barang],Inventory.UOM as Satuan From " & _
                             "TransData INNER JOIN [Detail TransData] ON (TransData.TransID = [Detail TransData].TransID) " & _
                             "INNER JOIN dbo.Inventory ON (dbo.[Detail TransData].NoItem = dbo.Inventory.NoItem) " & _
                             "Where (dbo.TransData.TypeTrans = 'AR') and TransData.TransID ='" & Params & "'", CNN, lckLockReadOnly
            mCall.FromTagActive = "INVENTORY"
End Select
Set mCall.FormData = RcPartner.DBRecordset
mCall.LookUp Me
   
End Sub


Private Sub cmdLink_Click(Index As Integer)
   OpenPartner Index, Text1(1)
End Sub

Private Sub cmdLinkDoc_Click(Index As Integer)

End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
Select Case AdReasonActiveDb
Case tmbAddNew:
     Text1(10).Text = IndexAuto
     cmdLink(0).Enabled = True
     cmdLink(1).Enabled = True
     cmdLink(3).Enabled = True
     cmdLink(4).Enabled = True
     DTPicker1(0).Enabled = True
     DTPicker1(1).Enabled = True
Case tmbEdit:
     cmdLink(0).Enabled = True
     cmdLink(1).Enabled = True
     cmdLink(3).Enabled = True
     cmdLink(4).Enabled = True
     DTPicker1(0).Enabled = True
     DTPicker1(1).Enabled = True
     Set DTPicker1(0).DataSource = MyDDE.ActiveRecordset
     Set DTPicker1(1).DataSource = MyDDE.ActiveRecordset
Case tmbCancel:
     cmdLink(0).Enabled = False
     cmdLink(1).Enabled = False
     cmdLink(3).Enabled = False
     cmdLink(4).Enabled = False
 Case tmbPrint:
            Dim aReport As New utility
            aReport.CallReportView "select * from memopotonganharga where [Memo ID]='" & Text1(10) & "'", "memo potongan harga.rpt", ReportPath, "Memo Motongan Harga"
            Set aReport = Nothing
End Select
Err.Clear
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
Select Case AdReasonActiveDb
Case tmbEdit:
Case tmbAddNew:
     Set DTPicker1(0).DataSource = MyDDE.ActiveRecordset
     Set DTPicker1(1).DataSource = MyDDE.ActiveRecordset
Case tmbSave:
     MyDDE.IsChildMemberReady = True
     simpan
     DTPicker1(0).Enabled = False
     DTPicker1(1).Enabled = False
     cmdLink(0).Enabled = False
     cmdLink(1).Enabled = False
     cmdLink(3).Enabled = False
     cmdLink(4).Enabled = False
Case tmbDelete
     MyDDE.PrepareDelete = "delete  from [memo potongan harga] where [memo id] = '" & Text1(10).Text & "'"
     DTPicker1(0).Enabled = False
     DTPicker1(1).Enabled = False
     cmdLink(0).Enabled = False
     cmdLink(1).Enabled = False
     cmdLink(3).Enabled = False
     cmdLink(4).Enabled = False
Case tmbCancel
     DTPicker1(0).Enabled = False
     DTPicker1(1).Enabled = False
End Select
End Sub
Function simpan()
MyDDE.PrepareAppend = "insert into [Memo Potongan Harga] ([memo id], [date memo], [partner ID],[No Faktur], [No Item], Discount, reason, lampiran, [Sales ID], [Marketing ID], [Doc No], Revisi, [Effective Date],ordered_by) values " & _
                    " ('" & Text1(10).Text & "', '" & Format(DTPicker1(0).Value, "yyyy-MM-dd") & "', N'" & MyDDE.GetFieldByName("Partner ID") & "', '" & Text1(1).Text & "', '" & Text1(2).Text & "', '" & Text1(3).Text & "', '" & Text1(4).Text & "', '" & Text1(5).Text & "', " & _
                    " '" & Text1(6).Text & "', '" & Text1(7).Text & "', '" & Text1(8).Text & "', '" & Text1(9).Text & "', '" & Format(DTPicker1(1).Value, "yyyy-MM-dd") & "','" & MainMenu.StatusBar1.Panels(1).Text & "')"
MyDDE.PrepareUpdate = "update [memo potongan harga] set [date memo] = '" & Format(DTPicker1(0).Value, "yyyy-MM-dd") & "', [partner id] = N'" & MyDDE.GetFieldByName("Partner ID") & "', [no faktur] = '" & Text1(1).Text & "', [no item] = '" & Text1(2).Text & "', " & _
                    " discount = '" & Text1(3).Text & "', reason = '" & Text1(4).Text & "', lampiran = '" & Text1(5).Text & "', [sales id] = '" & Text1(6).Text & "', [marketing ID] = '" & Text1(7).Text & "', [doc no] =  '" & Text1(8).Text & "', revisi = '" & Text1(9).Text & "', [effective date] = '" & Format(DTPicker1(1).Value, "yyyy-MM-dd") & "',ordered_by='" & MainMenu.StatusBar1.Panels(1).Text & "' where [memo id] = '" & Text1(10).Text & "'"
End Function

Private Function IndexAuto() As String
Dim Rc As New DBQuick
Dim TglSaiki As String
Dim Inom As Long
TglSaiki = Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2)
Rc.DBOpen "SELECT MAX(RIGHT([MEMO ID], 5)) AS MaxNom FROM [memo potongan harga] WHERE (GETDATE() = { fn NOW() })", CNN, lckLockReadOnly
With Rc
     If .DBRecordset.Recordcount <> 0 Then
        Inom = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + 1
     Else
        Inom = 1
     End If
     Select Case Len(Trim(Str(Inom)))
            Case 0: IndexAuto = "ME/" & TglSaiki & "-" & Trim(Str(Inom))
            Case 1: IndexAuto = "ME/" & TglSaiki & "-" & "0000" & Trim(Str(Inom))
            Case 2: IndexAuto = "ME/" & TglSaiki & "-" & "000" & Trim(Str(Inom))
            Case 3: IndexAuto = "ME/" & TglSaiki & "-" & "00" & Trim(Str(Inom))
            Case 4: IndexAuto = "ME/" & TglSaiki & "-" & "0" & Trim(Str(Inom))
     End Select
End With
End Function

Private Sub Form_Load()
On Error Resume Next

DTPicker1(0).Value = Now
DTPicker1(1).Value = Now
sql = "SELECT  dbo.[Memo Potongan Harga].[Memo ID], dbo.[Memo Potongan Harga].[Date Memo], dbo.[Memo Potongan Harga].[Partner ID]," & _
             " dbo.PartnerDB.ContactName, dbo.[Memo Potongan Harga].[No Faktur], dbo.[Memo Potongan Harga].[No Item], dbo.[Memo Potongan Harga].Discount," & _
             " dbo.[Memo Potongan Harga].reason, dbo.[Memo Potongan Harga].lampiran, dbo.[Memo Potongan Harga].[Sales ID]," & _
             " dbo.[Memo Potongan Harga].[Marketing ID], dbo.[Memo Potongan Harga].[Doc No], dbo.[Memo Potongan Harga].Revisi," & _
             " dbo.[Memo Potongan Harga].[Effective Date], dbo.PartnerDB.CompanyName, dbo.[Memo Potongan Harga].approved_by" & _
             " FROM dbo.[Memo Potongan Harga] INNER JOIN " & _
             " dbo.PartnerDB ON dbo.[Memo Potongan Harga].[Partner ID] = dbo.PartnerDB.PartnerID"

HiasFormManTell Picture2, Me
Set MyDDE.ActiveConnection = CNN
    MyDDE.BindFormTAG = "memo"
Set MyDDE.BindForm = Me
    MyDDE.PrepareQuery = sql
   
Set DTPicker1(0).DataSource = MyDDE.ActiveRecordset
Set DTPicker1(1).DataSource = MyDDE.ActiveRecordset
DTPicker1(0).Enabled = False
DTPicker1(1).Enabled = False
Err.Clear
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
If pRecordset.Recordcount <> 0 Then
Select Case UCase(TagForm):
       Case "CUSTOMER":
            MyDDE.GetFieldByName("Partner ID") = mCall.GetFieldByName(0)
            Text1(0) = mCall.GetFieldByName(1)
       Case "FAKTUR"
            MyDDE.GetFieldByName("purchaseID") = mCall.GetFieldByName(0)
            Text1(1) = mCall.GetFieldByName(0)
        Case "INVENTORY"
            'MyDDE.GetFieldByName("purchaseID") = mcall.GetFieldByName(0)
             Text1(2) = mCall.GetFieldByName(0)
End Select
End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
'Text1(Index).BackColor = &H79BCFF
End Sub

Private Sub Text1_LostFocus(Index As Integer)
'Text1(Index).BackColor = &HFCF1ED
End Sub
