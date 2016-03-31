VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmEmployess 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employes"
   ClientHeight    =   6975
   ClientLeft      =   1170
   ClientTop       =   5190
   ClientWidth     =   10035
   Icon            =   "frmEmployess.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   Tag             =   "Employee"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6405
      Left            =   0
      ScaleHeight     =   6405
      ScaleWidth      =   10035
      TabIndex        =   11
      Top             =   0
      Width           =   10035
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Mobile"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   6450
         MaxLength       =   15
         TabIndex        =   9
         Tag             =   "Partner"
         Top             =   2640
         Width           =   3255
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "HomePhone"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   6450
         MaxLength       =   15
         TabIndex        =   8
         Tag             =   "Partner"
         Top             =   2280
         Width           =   3255
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "PostalCode"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   6450
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Partner"
         Top             =   1920
         Width           =   3255
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "City"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   6450
         MaxLength       =   15
         TabIndex        =   6
         Tag             =   "Partner"
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Address"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   6450
         MaxLength       =   60
         TabIndex        =   5
         Tag             =   "Partner"
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "FullName"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   6450
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "Partner"
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "EmpID"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   6450
         MaxLength       =   15
         TabIndex        =   2
         Tag             =   "Partner"
         Top             =   120
         Width           =   3255
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Notes"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Index           =   8
         Left            =   6480
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Tag             =   "Partner"
         Top             =   3000
         Width           =   3255
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1485
         Top             =   1605
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   16777215
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployess.frx":6852
               Key             =   "Orang"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployess.frx":7426
               Key             =   "person1"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployess.frx":7D02
               Key             =   "person2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployess.frx":85DE
               Key             =   "TOP"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployess.frx":9432
               Key             =   "Dept"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   5880
         Left            =   105
         TabIndex        =   1
         Top             =   105
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   10372
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DgMaster 
         Height          =   135
         Left            =   3585
         TabIndex        =   12
         Tag             =   "Partner"
         Top             =   2940
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   238
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16777215
         HeadLines       =   2
         RowHeight       =   15
         RowDividerStyle =   6
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
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "EmpID"
            Caption         =   "Karyawan ID"
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
            DataField       =   "FullName"
            Caption         =   "Nama Lengkap"
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
            DataField       =   "JobPosition"
            Caption         =   "Jabatan"
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
            DataField       =   "BirthDate"
            Caption         =   "Tgl. Lahir"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd-MMM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Address"
            Caption         =   "Alamat"
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
            DataField       =   "City"
            Caption         =   "Kota"
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
            DataField       =   "PostalCode"
            Caption         =   "Kode Pos"
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
            DataField       =   "HomePhone"
            Caption         =   "Telp"
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
         BeginProperty Column08 
            DataField       =   "Mobile"
            Caption         =   "Mobile"
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
            MarqueeStyle    =   4
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "BirthDate"
         Height          =   345
         Left            =   6435
         TabIndex        =   4
         Tag             =   "Partner"
         Top             =   828
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   609
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
         Format          =   71630851
         CurrentDate     =   38357
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   13
         Left            =   5055
         TabIndex        =   21
         Top             =   2640
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   12
         Left            =   5055
         TabIndex        =   20
         Top             =   4065
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telp"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   6
         Left            =   5055
         TabIndex        =   19
         Top             =   2280
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pos Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   5
         Left            =   5055
         TabIndex        =   18
         Top             =   1890
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   4
         Left            =   5055
         TabIndex        =   17
         Top             =   1560
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   3
         Left            =   5055
         TabIndex        =   16
         Top             =   1230
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   1
         Left            =   5055
         TabIndex        =   15
         Top             =   525
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   0
         Left            =   5055
         TabIndex        =   14
         Top             =   210
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Birth"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   7
         Left            =   5055
         TabIndex        =   13
         Top             =   885
         Width           =   825
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   5040
         X2              =   6915
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   5040
         X2              =   6915
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   5040
         X2              =   6555
         Y1              =   1158
         Y2              =   1158
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   5040
         X2              =   6915
         Y1              =   1515
         Y2              =   1515
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   5040
         X2              =   6915
         Y1              =   1875
         Y2              =   1875
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   5040
         X2              =   6915
         Y1              =   2235
         Y2              =   2235
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   5040
         X2              =   7395
         Y1              =   2595
         Y2              =   2595
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   5040
         X2              =   7410
         Y1              =   2955
         Y2              =   2955
      End
      Begin VB.Line Line1 
         Index           =   9
         X1              =   5040
         X2              =   6915
         Y1              =   4515
         Y2              =   4515
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   6405
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmEmployess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private mVarnode As Nodes
Private mVarNodeNode As Node
Private MEdit As Boolean
Private MyTrans As New clsTransaksi
Private mVarParentNode As String
Private mVarIndex As String
Private OrderJabatan As Boolean

'Private Sub cmdLink_Click(Index As Integer)
'OpenPartner Index
'End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
'HiasForm Picture1, Me

MyDDE.SetPermissions = aksess.MayDo("Employee")

HiasFormManTell Picture2, Me
With MyDDE
    .EditModeReplace = False
    .BindFormTAG = "Partner"
    Set .BindForm = frmEmployess
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT     Employees.EmpID AS [Employee ID], Employees.FullName AS [Full Name], Employees.BirthDate, Employees.Address, Employees.City,                        Employees.PostalCode AS [Postal Code], Employees.HomePhone AS Phone, Employees.Mobile AS Mobile, Employees.Photo AS Notes,                        Employees.Notes AS Expr1, Employees.ReportsTo FROM         [Tabel Departemen] INNER JOIN                      Employees ON [Tabel Departemen].[Kode Dep] = Employees.[Kode Dep] WHERE     ([Tabel Departemen].[Kode Dep] = 0) ORDER BY Employees.EmpID"
End With
Set mVarnode = TreeView1.Nodes
TreeView1.Indentation = 300 '19 * Screen.TwipsPerPixelX
LoadTree
'TreeView1.Indentation = 1000
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MyDDE.CheckRecordPendinged = True Then
   ScanKey vbKeyF5, 0, MyDDE
   If MyDDE.IsSucces = True Then
      Cancel = False
      MyDDE.ClearRecordset
   Else
      Cancel = True
   End If
Else
   MyDDE.ClearRecordset
End If
End Sub

Private Sub Form_Resize()

'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmEmployess = Nothing
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
On Error GoTo 1
Select Case TagForm
       Case "MASTER JABATAN":
            MyDDE.GetFieldByName("Nama Jabatan") = mCall.GetFieldByName(0)
            MyDDE.GetFieldByName("Kode Jabatan") = mCall.GetFieldByName(1)
            
       Case "MASTER DEPARTEMEN":
            MyDDE.GetFieldByName("DEPARTEMENT") = mCall.GetFieldByName(0)
       Case "PENANGGUNG JAWAB":
            MyDDE.GetFieldByName("ReportsTO") = mCall.GetFieldByName(0)
End Select
Exit Sub
1:
MessageBox Err.Description, "frmemployees_mcall_rowcolchange" & Err.Number, msgOkOnly, msgExclamation
End Sub

Public Sub mnEdit_Click()
TreeView1.StartLabelEdit
''TreeView1.LabelEdit = tvwAutomatic
'mVarNode.Item(1).Text = mVarNode.Item(1).Selected
End Sub

Public Sub mnHapus_Click()
Dim I As Integer
If IsHasChild = False Then
   I = MessageBox("Anda yakin untuk menghapus data departemen/divisi?", "Penghapusan", msgYesNo)
   If I = 1 Then AddNode "delete"
Else
   MessageBox "Data departemen/divisi tidak bisa dihapus karena masih dipakai.", "Penghapusan", msgOkOnly
End If
End Sub

Public Sub mnJabat_Click()
OrderJabatan = True
AddNode "TAMBAH", OrderJabatan
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 1
Select Case AdReasonActiveDb
       Case tmbAddNew:
            'mVarDataDc = True
            txtBox(0).SetFocus
       Case tmbEdit:
            txtBox(0).Enabled = False
            'mVarDataDc = True
            txtBox(1).SetFocus
       Case tmbSave:
            'AddNode "Tambah"
            AddUser "tambah"
       Case tmbPrint:
            CallRPTReport "Employee.rpt"
       Case Else: 'mVarDataDc = False
End Select
Exit Sub
1:
MessageBox Err.Description, "frmemployees_mydde_afterprepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 2
Dim mDel As New clsDelete
Select Case AdReasonActiveDb
       Case tmbEdit:
             If IsHasChildren = False Then
                MEdit = True
             Else
                MyDDE.CancelTrans = True
                MessageBox "Top level struktur organisasi/departemen tidak dapat diisi data karyawan."
             End If
'            If IsNodeOk(Val(Replace(mVarParentNode, "ANAK", ""))) = True Then
'               mEdit = True
'            Else
'               MyDDE.CancelTrans = True
'               MessageBox "Top level struktur organisasi/departemen tidak dapat diisi data karyawan."
'            End If
       Case tmbAddNew:
            If IsNodeOk(Val(Replace(mVarParentNode, "ANAK", "")), True) = True Then
               MEdit = True
            Else
               MyDDE.CancelTrans = True
               MessageBox "Top level struktur organisasi/departemen tidak dapat diisi data karyawan." & vbCrLf & "Dan jika jabatan sudah berisi karyawan tidak bisa ditambah."
            End If
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               If mDel.CekDelete(txtBox(0), reDelMasterCurency) = False Then
                  MyDDE.IsChildMemberReady = True
                  PrepareQuery
                  MEdit = False
               Else
                  MyDDE.CancelTrans = True
                  MessageBox "Record (" & txtBox(0) & ") Sedang Dipakai Transaksi Lain." & vbCrLf & "Record Tidak Bisa DiHapus.", "Peringatan", msgOkOnly
                  MyDDE.IsChildMemberReady = False
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               MyDDE.GetFieldByName("BirthDate") = DTPicker1.Value
               PrepareQuery
               MEdit = False
            Else
               MyDDE.IsChildMemberReady = False
            End If
      Case tmbCancel: MEdit = False
End Select
Set mDel = Nothing
'cmdLink(0).Enabled = mEdit
'cmdLink(2).Enabled = mEdit
2:
MessageBox Err.Description, "frmemployees_mydde_prepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

Private Sub TreeView1_AfterLabelEdit(Cancel As Integer, NewString As String)
mVarnode.Item(TreeView1.SelectedItem.Index).Text = NewString
SendDataToServer (" Update [Tabel Departemen]" & _
                  " Set [Nama Dep] = N'" & NewString & "'" & _
                  " WHERE  ([Kode Dep] = " & Replace(TreeView1.SelectedItem.Key, "ANAK", "") & ")")

End Sub

Public Sub mnTdep_Click()
TreeView1.StartLabelEdit
OrderJabatan = False
mVarParentNode = "1"
AddNode "TAMBAH", OrderJabatan
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)
'TreeView1.LabelEdit = tvwAutomatic
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
Dim I As Integer
On Error Resume Next
Set mVarNodeNode = Node
mVarParentNode = Node.Key
mVarIndex = Node.Index
OpenDB Val(Replace(mVarParentNode, "ANAK", ""))

Err.Clear
End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Select Case Button
       Case vbLeftButton:   TreeView1.Tag = Button
       Case vbRightButton:
            If TreeView1.Tag = "1" Then
               Set TreeView1.SelectedItem = TreeView1.HitTest(x, Y)
               If Not TreeView1.SelectedItem Is Nothing Then PopupMenu MainMenu.MnNodes, , , , MainMenu.mnTdep
            End If
End Select
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub PrepareQuery()
With MyDDE
    .PrepareAppend = " INSERT INTO Employees" & _
                     " (EmpID, FullName, [Kode Dep], BirthDate, Address, City, PostalCode, HomePhone, Mobile, Notes)" & _
                     " VALUES     (N'" & txtBox(0) & "', N'" & ValidString(txtBox(1)) & "'," & Replace(mVarParentNode, "ANAK", "") & ", CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), N'" & ValidString(txtBox(3)) & "', N'" & ValidString(txtBox(4)) & "', N'" & ValidString(txtBox(5)) & "', N'" & ValidString(txtBox(6)) & "', N'" & ValidString(txtBox(7)) & "', N'" & ValidString(txtBox(8)) & "')"
                     
    .PrepareUpdate = " UPDATE [Employees] Set  [FullName] = N'" & ValidString(txtBox(1)) & "',[Kode Dep]=" & Replace(mVarParentNode, "ANAK", "") & ",BirthDate = CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3),Address = N'" & ValidString(txtBox(3)) & "',City = N'" & ValidString(txtBox(4)) & "',PostalCode = N'" & ValidString(txtBox(5)) & "',HomePhone = N'" & ValidString(txtBox(6)) & "',Mobile = N'" & ValidString(txtBox(7)) & "',Notes = N'" & ValidString(txtBox(8)) & "' WHERE     (EmpID = N'" & ValidString(txtBox(0)) & "')"
    
    .PrepareDelete = " DELETE FROM [Employees] WHERE   (EmpID = N'" & ValidString(txtBox(0)) & "') "
End With
End Sub

Private Sub txtBox_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 6 Or Index = 7 Then
   ValidNum KeyAscii
End If
End Sub
Private Function OpenPartner(ByVal Index As Integer) As Boolean
Dim RcPartner As New DBQuick
Set mCall = New frmCaller
Select Case Index
       Case 0:
            RcPartner.DBOpen " SELECT     [Nama Dep] AS [Nama Jabatan], [Kode Dep] AS [Kode Jabatan] FROM         [Tabel Departemen] WHERE     (Type = 1) ORDER BY [Nama Dep]", CNN, lckLockReadOnly
            mCall.FromTagActive = "MASTER JABATAN"
       Case 1:
            RcPartner.DBOpen " SELECT     [Nama Dep] AS Departemen, [Kode Dep] AS [Kode departemen] FROM         [Tabel Departemen] ORDER BY [Nama Dep]", CNN, lckLockReadOnly
            mCall.FromTagActive = "MASTER DEPARTEMEN"
            'mCall.txtCari = NoVoucher(1)
       Case 2:
            RcPartner.DBOpen " SELECT     [Tabel Jabatan].[Kode Jabat] AS [Kode Jabatan], [Tabel Jabatan].[Nama Jabatan], [Tabel Departemen].[Nama Dep] AS [Nama Departemen] FROM         [Tabel Jabatan] INNER JOIN                       [Tabel Departemen] ON [Tabel Jabatan].[Kode Dep] = [Tabel Departemen].[Kode Dep]", CNN, lckLockReadOnly
            mCall.FromTagActive = "PENANGGUNG JAWAB"
End Select
If RcPartner.Recordcount <> 0 Then
    Set mCall.FormData = RcPartner.DBRecordset
    mCall.LookUp Me
    'If MyDDE.ChildRecordset.Recordcount <> 0 Then
       
'    If FindOwnRecordset(MyDDE.ChildRecordset, "[Kode Akun] = '" & MyDDE.ChildRecordset.Fields(0) & "'") = True Then
'       MessageBox "Record -> " & MyDDE.ChildRecordset.Fields(0) & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
'       MyDDE.ChildRecordset.CancelBatch adAffectCurrent
'       If MyDDE.ChildRecordset.Recordcount >= 1 Then MyDDE.ChildRecordset.MoveLast
'
'       DGPurchase.SetFocus
'    End If
    'End If
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly
   OpenPartner = True
End If
RcPartner.CloseDB
Set mCall = Nothing
End Function

Private Sub AddNode(ByVal Tipical As String, Optional ByVal Departement_OR_JABATAN As Boolean = False)
On Error GoTo 1
Dim I As Integer
Dim StrKey As String
Dim StrParent As String
Dim StrText As String
If Departement_OR_JABATAN = False Then I = 0 Else I = 1
Select Case UCase(Tipical)
       Case "TAMBAH":
            
            If mVarnode.Count = 0 Then
               StrKey = IndexKu & "ANAK"
               StrParent = "TOP"
               StrText = "Tambah Departemen"
               With mVarnode.Add(, , StrKey, "Tambah Departemen", "TOP")
                    .Expanded = True
                    .Bold = True
               End With
             Else
               StrKey = IndexKu & "ANAK"
               StrParent = Replace(mVarParentNode, "ANAK", "")
               If StrParent = "1" Then
                  StrText = "Tambah Departemen" & mVarnode.Count
                  With mVarnode.Add(1 & "ANAK", tvwChild, StrKey, StrText, "Dept")
                       .Expanded = True
                  End With
               Else
                  StrText = "Tambah Jabatan" & mVarnode.Count
                  With mVarnode.Add(StrParent & "ANAK", tvwChild, StrKey, StrText, "Orang")
                       .Expanded = True
                  End With
               End If
             End If
             SendDataToServer (" INSERT INTO [Tabel Departemen] " & _
                               " ([Kode Dep], [Nama Dep], ReportsTo,[Type])" & _
                               " VALUES  (" & Replace(StrKey, "ANAK", "") & ", N'" & StrText & "', N'" & StrParent & "'," & I & ")")
             
             mVarnode.Item(mVarnode.Count).Selected = True
             TreeView1.StartLabelEdit
       Case "DELETE":
            If IsNodeReady = True Then
               mVarnode.Remove TreeView1.SelectedItem.Index
               SendDataToServer (" Delete from [Tabel Departemen] where [Kode Dep] =" & Val(mVarParentNode))
            Else
               MessageBox "Node user belum dipilih. Silahkan diulangi."
            End If
End Select
Exit Sub
1:
MessageBox Err.Description, "frmemployess:addnode" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub LoadTree()
Dim rcNode As New DBQuick
Dim I As Integer
Dim Avdata As Variant
Dim StrImage As String
rcNode.DBOpen "SELECT    * from [Tabel Departemen] order by [Kode Dep]", CNN
With rcNode
     If .DBRecordset.Recordcount <> 0 Then
        Avdata = .DBRecordset.Getrows(.DBRecordset.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            If Avdata(2, I) = "TOP" Then
               With mVarnode.Add(, , 1 & "ANAK", Avdata(1, I), "TOP")
                    .Expanded = True
                    .Bold = True
               End With
            Else
               If Avdata(2, I) = 1 Then
                  StrImage = "Dept"
               Else
                  StrImage = "Orang"
               End If
               With mVarnode.Add(Avdata(2, I) & "ANAK", tvwChild, Avdata(0, I) & "ANAK", Avdata(1, I) & CarikodeData(Avdata(0, I)), StrImage)
                    .Expanded = True
               End With
            End If
        Next I
     Else
     End If
End With
End Sub

Private Function IsHasChild() As Boolean
On Error GoTo Hell
If mVarNodeNode.Child Is Nothing Then
   IsHasChild = False
Else
   IsHasChild = True
End If
Exit Function
Hell:
   IsHasChild = False
End Function

Private Function IndexKu(Optional ByVal Tipical As Boolean = False) As Long
On Error GoTo 2

Dim Rc As New DBQuick
Dim I As Long
If Tipical = False Then
   Rc.DBOpen "SELECT     MAX([Kode Dep]) AS Expr1 FROM         [Tabel Departemen]", CNN
Else
   Rc.DBOpen "SELECT     MAX([Kode Jabat]) AS Expr1 FROM         [Tabel Jabatan]", CNN
End If
With Rc
     If .Recordcount <> 0 Then
        IndexKu = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + 1
     Else
        IndexKu = 1
     End If
End With
Exit Function
2:
MessageBox Err.Description, "frmemployess:indexku" & Err.Number, msgOkOnly, msgExclamation
End Function

Private Function CarikodeData(ByVal Param As Long) As String
Dim Rc As New DBQuick
Rc.DBOpen "SELECT     ' (' + EmpID + '-' + FullName + ')' AS Nama FROM         Employees WHERE     ([Kode Dep] = " & Param & ")", CNN
'messagebox Rc.DBRecordset.Source
With Rc
     If .Recordcount <> 0 Then CarikodeData = IIf(Not IsNull(.Fields(0)), .Fields(0), "xxx")
End With
End Function

Private Function IsNodeOk(ByVal Param As Long, Optional ByVal TipicalTambah_OR_Edit As Boolean = False) As Boolean
Dim Rc As New DBQuick
Rc.DBOpen "SELECT     [Tabel Departemen].Type, Employees.EmpID FROM         [Tabel Departemen] LEFT OUTER JOIN                       Employees ON [Tabel Departemen].[Kode Dep] = Employees.[Kode Jabat] WHERE     ([Tabel Departemen].[Kode Dep] = " & Param & ") and ([Tabel Departemen].Type)=1", CNN
With Rc.DBRecordset
     Select Case .Recordcount
            Case 0:
                 IsNodeOk = False
            Case 1:
                 If IsNull(.Fields(1)) Then
                    'If TipicalTambah_OR_Edit = True Then
                       IsNodeOk = TipicalTambah_OR_Edit
                    'Else
                    '   IsNodeOk = False
                    'End If
                 Else
                    IsNodeOk = False
                 End If
            Case Is >= 2: IsNodeOk = False
            Case Else
                 IsNodeOk = False
     End Select
End With
End Function

Private Sub AddUser(ByVal Tipical As String)
Dim StrFirstText As String
Select Case UCase(Tipical)
       Case "TAMBAH"
            StrFirstText = mVarNodeNode.Text
            mVarnode.Item(Val(mVarIndex)).Text = StrFirstText & " (" & txtBox(0) & "-" & txtBox(1) & ")"

End Select
End Sub

Private Sub OpenDB(ByVal Param As Long)
MyDDE.PrepareQuery = "SELECT     Employees.EmpID, Employees.FullName, Employees.BirthDate, Employees.Address, Employees.City, Employees.PostalCode,                        Employees.HomePhone, Employees.Mobile, Employees.Photo, Employees.Notes, Employees.ReportsTo FROM         [Tabel Departemen] INNER JOIN                       Employees ON [Tabel Departemen].[Kode Dep] = Employees.[Kode dep] WHERE     (Employees.[Kode Dep] = " & Param & ") ORDER BY Employees.EmpID"
End Sub

Private Function IsNodeReady() As Boolean
Dim I As Integer
On Error GoTo Hell
   I = mVarNodeNode.Index
   IsNodeReady = True
Hell:
End Function
'
Private Function IsHasChildren() As Boolean
On Error GoTo Hell
With mVarNodeNode
     If .Children > 0 Then IsHasChildren = True
End With
Exit Function
Hell:
    Err.Clear
End Function
