VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{11D78E78-0CB5-48CD-ADB4-348FD684EE87}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmPurchaseOffer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Offer"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10080
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPurchaseOffer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   10080
   Begin SemeruDC.SemeruOleDC SDC 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   6435
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   1005
      BindFormTAG     =   "POffer"
      InitControlSet  =   1
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   120
      ScaleHeight     =   6225
      ScaleWidth      =   9825
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   5535
         Left            =   120
         ScaleHeight     =   5505
         ScaleWidth      =   9585
         TabIndex        =   1
         Top             =   600
         Width           =   9615
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "PurchaseID"
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   6
            Tag             =   "PO"
            Top             =   240
            Width           =   2370
         End
         Begin VB.CommandButton cmdLink 
            Height          =   330
            Index           =   0
            Left            =   8895
            Picture         =   "FrmPurchaseOffer.frx":6852
            Style           =   1  'Graphical
            TabIndex        =   5
            Tag             =   "POffer"
            Top             =   225
            Width           =   330
         End
         Begin VB.TextBox lblSpl 
            Appearance      =   0  'Flat
            DataField       =   "CompanyName"
            Height          =   330
            Left            =   6480
            TabIndex        =   4
            Tag             =   "POffer"
            Top             =   225
            Width           =   2385
         End
         Begin MSDataGridLib.DataGrid GridPO 
            Height          =   4335
            Left            =   120
            TabIndex        =   3
            Top             =   1080
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   7646
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
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
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "DatePurchase"
            Height          =   315
            Left            =   1560
            TabIndex        =   7
            Tag             =   "POffer"
            Top             =   585
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   62324739
            CurrentDate     =   38272
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            DataField       =   "DatePurchase"
            Height          =   315
            Left            =   6510
            TabIndex        =   12
            Tag             =   "POffer"
            Top             =   645
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   62324739
            CurrentDate     =   38272
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Require Date"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   2
            Left            =   5160
            TabIndex        =   13
            Top             =   720
            Width           =   945
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   5160
            X2              =   6585
            Y1              =   945
            Y2              =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Order Date"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   9
            Left            =   240
            TabIndex        =   11
            Top             =   660
            Width           =   915
         End
         Begin VB.Label lblSupplier 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Partner ID"
            DataField       =   "Address"
            ForeColor       =   &H80000005&
            Height          =   210
            Left            =   4680
            TabIndex        =   10
            Tag             =   "PO"
            Top             =   1080
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   4
            Left            =   5160
            TabIndex        =   9
            Top             =   285
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   8
            Top             =   300
            Width           =   165
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   240
            X2              =   1665
            Y1              =   540
            Y2              =   540
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   240
            X2              =   1665
            Y1              =   885
            Y2              =   885
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   5160
            X2              =   6585
            Y1              =   540
            Y2              =   540
         End
      End
   End
End
Attribute VB_Name = "FrmPurchaseOffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsDetail As New DBQuick
Dim IDGen As New IDGenerator
Private MyData As New clsTransaksi

Private Sub DataGrid1_Click()

End Sub

Private Sub Form_Load()
   HiasForm Picture1, Me
   CenterForm Picture2, Me
   Set SDC.BindForm = Me
   Set SDC.ActiveConnection = CNN
   SDC.PrepareQuery = "Select POffer.*,partnerDB.companyName from POffer inner join partnerDB on POffer.partnerID = partnerDB.partnerID where PartnerDB.PartnerID like 'SP%'"
   'Load Data Dept
   
End Sub

Private Sub SDC_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo xErr
   With SDC.ActiveRecordset
      txtBox.Text = .Fields("ID").Value
      lblSpl.Text = .Fields("CompanyName").Value
      lblSupplier.Caption = .Fields("PartnerId").Value
      DTPicker1.Value = .Fields("DateOffer").Value
      DTPicker2.Value = .Fields("Require Date").Value
   End With
   loadDetail
Exit Sub
xErr:
   Err.Clear
End Sub

Private Sub loadDetail()
   rsDetail.DBOpen "Select Inventory.itemName,inventory.UOM ,[Detail POffer].* from [Detail POffer] inner join Inventory on [Detail POffer].prodID = inventory.NoItem where ID='" & txtBox.Text & "'", CNN
   Set SDC.ChildRecordset = rsDetail.DBRecordset
   Set GridPO.DataSource = rsDetail.DBRecordset
   GridLayout
End Sub
   
Private Sub GridLayout()
   With GridPO
      .Columns(0).Caption = "Nama Item"
      .Columns(0).Width = 3000
      .Columns(1).Caption = "Satuan"
      .Columns(1).Width = 1000
      .Columns(2).Visible = False
      .Columns(3).Visible = False
      .Columns(4).Visible = False
      .Columns(6).Caption = "Keterangan"
      .Columns(6).Width = 3500
      .Columns(7).Visible = False
      .Columns(8).Visible = False
   End With
End Sub


Private Sub SDC_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
   Select Case AdReasonActiveDb
      Case tmbAddNew:
         txtBox.Text = IDGen.GetID("OF")
         DTPicker1.Value = Now
         DTPicker2.Value = Now
      
      Case tmbDetail:
            If SDC.CheckEmptyControl = False Then
               If MyData.CheckGridKosong(SDC.ChildRecordset, "Qty") = True Then
                   SDC.CancelTrans = True
                   MessageBox "Data transaksi belum lengkap." & "Silahkan dicek kembali.", "Peringatan", msgOkOnly
               End If
            Else
               SDC.CancelTrans = mFirstCaller
            End If
      Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If CekGridKosong = False And MyDDE.ChildRecordset.Recordcount <> 0 Then
                  MyDDE.IsChildMemberReady = True
                  MyDDE.GetFieldByName("DatePurchase") = DTPicker1.Value
                  'MyDDE.GetFieldByName("DatePurchase") = DTPicker1.Value
'                  'PrepareQuery
               Else
                  MyDDE.IsChildMemberReady = False
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
      
      Case tmbDetail:
      
   End Select
End Sub
