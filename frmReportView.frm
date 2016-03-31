VERSION 5.00
Object = "{FB992564-9055-42B5-B433-FEA84CEA93C4}#11.0#0"; "crviewer.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReportView 
   Caption         =   "Laporan"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   9765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReportView.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   9765
   Tag             =   "RPT"
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   " Tools "
      Height          =   5235
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   810
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   4020
         Left            =   150
         TabIndex        =   27
         Top             =   315
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   7091
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Document"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Printer"
               ImageIndex      =   13
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Send Document To Printer"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Enabled         =   0   'False
                     Object.Visible         =   0   'False
                     Text            =   "Papersize Option"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Printer Properties"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Export"
               ImageIndex      =   15
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Refresh"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Group Tree"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Zoom"
               ImageIndex      =   6
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   11
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Paper Fit"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "100%"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "90%"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "80%"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "70%"
                  EndProperty
                  BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "60%"
                  EndProperty
                  BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "50%"
                  EndProperty
                  BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "40%"
                  EndProperty
                  BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "30%"
                  EndProperty
                  BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "20%"
                  EndProperty
                  BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "10%"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Search"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Top Page Report"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Next Page Report"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Previous Page Report"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Bottom Page Report"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Quit Report Application"
               ImageIndex      =   12
            EndProperty
         EndProperty
      End
      Begin VB.Label LblCap 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   60
         TabIndex        =   30
         Top             =   7440
         Width           =   1125
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Page "
         Height          =   195
         Left            =   195
         TabIndex        =   29
         Top             =   4560
         Width           =   420
      End
      Begin VB.Label LblPages 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0/0"
         Height          =   195
         Left            =   255
         TabIndex        =   28
         Top             =   4800
         Width           =   240
      End
   End
   Begin CrystalActiveXReportViewerLib11Ctl.CrystalActiveXReportViewer CRviewReport 
      Height          =   2895
      Left            =   2520
      TabIndex        =   25
      Top             =   960
      Width           =   2895
      _cx             =   5106
      _cy             =   5106
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1057
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EAAF6F&
      Enabled         =   0   'False
      Height          =   3255
      Left            =   2820
      ScaleHeight     =   3195
      ScaleWidth      =   6105
      TabIndex        =   5
      Top             =   4365
      Visible         =   0   'False
      Width           =   6165
      Begin VB.ComboBox cboPaper 
         BackColor       =   &H00FCF1ED&
         Height          =   315
         Index           =   3
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   255
         Width           =   4020
      End
      Begin VB.TextBox TxtMargin 
         BackColor       =   &H00FCF1ED&
         Height          =   315
         Index           =   3
         Left            =   4500
         TabIndex        =   23
         Top             =   2205
         Width           =   915
      End
      Begin VB.TextBox TxtMargin 
         BackColor       =   &H00FCF1ED&
         Height          =   315
         Index           =   2
         Left            =   1395
         TabIndex        =   21
         Top             =   2205
         Width           =   915
      End
      Begin VB.TextBox TxtMargin 
         BackColor       =   &H00FCF1ED&
         Height          =   315
         Index           =   1
         Left            =   4500
         TabIndex        =   19
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox TxtMargin 
         BackColor       =   &H00FCF1ED&
         Height          =   315
         Index           =   0
         Left            =   1395
         TabIndex        =   17
         Top             =   1740
         Width           =   915
      End
      Begin VB.CommandButton CmdOk 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Print"
         Height          =   405
         Index           =   4
         Left            =   2175
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2640
         Width           =   1260
      End
      Begin VB.CommandButton CmdOk 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Quit"
         Height          =   405
         Index           =   2
         Left            =   4710
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2640
         Width           =   1260
      End
      Begin VB.CommandButton CmdOk 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Select Option"
         Height          =   405
         Index           =   3
         Left            =   3450
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2640
         Width           =   1260
      End
      Begin VB.ComboBox cboPaper 
         BackColor       =   &H00FCF1ED&
         Height          =   315
         Index           =   2
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   585
         Width           =   4020
      End
      Begin VB.ComboBox cboPaper 
         BackColor       =   &H00FCF1ED&
         Height          =   315
         Index           =   1
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1245
         Width           =   4020
      End
      Begin VB.ComboBox cboPaper 
         BackColor       =   &H00FCF1ED&
         Height          =   315
         Index           =   0
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   915
         Width           =   4020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left  Margin                             mm"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FCF1ED&
         Height          =   195
         Index           =   7
         Left            =   3195
         TabIndex        =   22
         Top             =   2265
         Width           =   2640
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Right Margin                        mm"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FCF1ED&
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   20
         Top             =   2265
         Width           =   2490
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bottom Margin                        mm"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FCF1ED&
         Height          =   195
         Index           =   5
         Left            =   3195
         TabIndex        =   18
         Top             =   1800
         Width           =   2670
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Top Margin                           mm"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FCF1ED&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   2490
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FCF1ED&
         Height          =   1485
         Left            =   150
         Top             =   150
         Width           =   5820
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FCF1ED&
         Height          =   915
         Left            =   150
         Top             =   1665
         Width           =   5820
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Defaut Printer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FCF1ED&
         Height          =   195
         Index           =   3
         Left            =   255
         TabIndex        =   14
         Top             =   315
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paper Source"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FCF1ED&
         Height          =   195
         Index           =   2
         Left            =   255
         TabIndex        =   10
         Top             =   645
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paper Orientation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FCF1ED&
         Height          =   195
         Index           =   1
         Left            =   255
         TabIndex        =   8
         Top             =   1305
         Width           =   1500
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Papersize"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FCF1ED&
         Height          =   195
         Index           =   0
         Left            =   255
         TabIndex        =   6
         Top             =   975
         Width           =   825
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FCF1ED&
         Height          =   3015
         Left            =   90
         Top             =   105
         Width           =   5940
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EAAF6F&
      Enabled         =   0   'False
      Height          =   1395
      Left            =   2805
      ScaleHeight     =   1335
      ScaleWidth      =   5190
      TabIndex        =   3
      Top             =   2895
      Visible         =   0   'False
      Width           =   5250
      Begin VB.CommandButton CmdOk 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Quit"
         Height          =   405
         Index           =   1
         Left            =   3765
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   705
         Width           =   1260
      End
      Begin VB.CommandButton CmdOk 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Find Next"
         Height          =   405
         Index           =   0
         Left            =   2505
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   705
         Width           =   1260
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FCF1ED&
         Height          =   315
         Left            =   1410
         TabIndex        =   0
         Top             =   330
         Width           =   3615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FCF1ED&
         Height          =   1185
         Left            =   90
         Top             =   105
         Width           =   5025
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Text Criteria"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FCF1ED&
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   390
         Width           =   1065
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1470
      Top             =   510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportView.frx":6852
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportView.frx":D0B4
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportView.frx":13916
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportView.frx":1A178
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportView.frx":209DA
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportView.frx":2723C
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportView.frx":2DA9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportView.frx":34300
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportView.frx":3AB62
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportView.frx":413C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportView.frx":47C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportView.frx":4E488
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmReportView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Rc As New Recordset, Rs1   As New Recordset, Rs2   As New Recordset, Rs3   As New Recordset, Rs4 As New Recordset

Private repName1 As String, repName2 As String, repName3 As String, repName4 As String

Private mStr As String

Private RcSubReport As New Recordset

Private m_Application As New CRAXDRT.Application

Public m_Report As CRAXDRT.Report

Private mTitle As String

Private mLoad As Boolean

Private Declare Function OpenPrinter _
                Lib "winspool.drv" _
                Alias "OpenPrinterA" (ByVal pPrinterName As String, _
                                      phPrinter As Long, _
                                      ByVal pDefault As Long) As Long

Private Declare Function ClosePrinter _
                Lib "winspool.drv" (ByVal hPrinter As Long) As Long

Private Declare Function DeviceCapabilities _
                Lib "winspool.drv" _
                Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, _
                                             ByVal lpPort As String, _
                                             ByVal iIndex As Long, _
                                             lpOutput As Any, _
                                             ByVal dev As Long) As Long

Private Const DC_BINS = 6

Private Const DC_BINNAMES = 12

Private KethokMas As Boolean
Dim Myrpt As New utility

Public Property Let QuerySource(ByVal vNewValue As String)
  On Error Resume Next
  Rc.CursorLocation = adUseClient
  Rc.Open vNewValue, CNN.ConnectionString, adOpenForwardOnly, adLockReadOnly, adCmdText
  Set Rc.ActiveConnection = Nothing
  'messagebox Rc.Recordcount
  Err.Clear
End Property

Public Property Let QuerySubReportSource(ByVal QuerySource As String, ByVal ReportName As String)
  On Error Resume Next
  RcSubReport.CursorLocation = adUseClient
  RcSubReport.Open QuerySource, CNN.ConnectionString, adOpenForwardOnly, adLockReadOnly, adCmdText
  Set RcSubReport.ActiveConnection = Nothing
  m_Report.OpenSubreport(ReportName).Database.SetDataSource RcSubReport

  'messagebox Err.Description
  Err.Clear
End Property

Public Sub SubReport1(ByVal SelectQuery As String, _
                      ByVal ReportName As String)
  Set Rs1 = Myrpt.OpenDB(SelectQuery)
  repName1 = ReportName
End Sub

Public Sub CallReport(ByVal SelectQuery As String, _
                      ByVal ReportName As String, _
                      ByVal ReportFilePath As String, _
                      ByVal ReportTitle As String, _
                      Optional ReportParams As String)
  'Dim Myrpt As New Utility
  mStr = SelectQuery
  Set Rc = Myrpt.OpenDB(SelectQuery)
  'Debug.Print SelectQuery
  ReportFilePath = ReportPath

  With Rc
    Set m_Report = m_Application.OpenReport(ReportFilePath & "\" & ReportName)

    If ReportTitle <> "" Then m_Report.ReportTitle = ReportTitle
    If ReportParams <> "" Then
      m_Report.ParameterFields.GetItemByName("Terbilang").AddCurrentValue ReportParams
    End If
  End With

  'Set Myrpt = Nothing
End Sub

Public Sub SubReport2(ByVal SelectQuery As String, _
                      ByVal ReportName As String)
  Set Rs2 = Myrpt.OpenDB(SelectQuery)
  repName2 = ReportName
End Sub

Public Sub SubReport3(ByVal SelectQuery As String, _
                      ByVal ReportName As String)
  Set Rs3 = Myrpt.OpenDB(SelectQuery)
  repName3 = ReportName
End Sub

Public Sub SubReport4(ByVal SelectQuery As String, _
                      ByVal ReportName As String)
  Set Rs4 = Myrpt.OpenDB(SelectQuery)
  repName4 = ReportName
End Sub

Public Property Let SubReportName(ByVal vNewValue As String)
  On Error GoTo Hell
  'Dim mRpt As New CRAXDRT.Report
  'Dim mApp As New CRAXDRT.Application
  m_Report.OpenSubreport (App.Path & "\Report\" & vNewValue)
  Exit Property
Hell:


  If Err.Number = -2147189547 Then
    '       'messagebox "Tidak bisa menampilkan report Lagi. Silahkan ditutup report yang sudah ada. Kemudian silahkan dibuka kembali report yang gagal ditampilkan.", vbCritical, "Report Warning"
    '       TutupReport
    '       Set mRpt = mApp.OpenReport(App.Path & "\Report\" & vNewValue)
    '       Set m_Report = Nothing
    '       Set m_Report = mRpt
  Else
    MessageBox "PROC_ReportName_ERROR" & vbCrLf & vNewValue & vbCrLf & Err.Description, vbExclamation, "Report Warning"
  End If

  Err.Clear
End Property

Public Property Let ReportName(ByVal vNewValue As String)
  On Error GoTo Hell
  Dim mRpt As New CRAXDRT.Report
  Dim mApp As New CRAXDRT.Application
  Set m_Report = m_Application.OpenReport(ReportPath & "\" & vNewValue)
  Exit Property
Hell:


  If Err.Number = -2147189547 Then
    'messagebox "Tidak bisa menampilkan report Lagi. Silahkan ditutup report yang sudah ada. Kemudian silahkan dibuka kembali report yang gagal ditampilkan.", vbCritical, "Report Warning"
    TutupReport
    Set mRpt = mApp.OpenReport(ReportPath & "\" & vNewValue)
    Set m_Report = Nothing
    Set m_Report = mRpt
  Else
    MessageBox "PROC_ReportName_ERROR" & vbCrLf & vNewValue & vbCrLf & Err.Description, vbCritical, "Report Warning"
  End If

  Err.Clear
End Property

Public Property Let PaperSizePrint(ByVal vNewValue As Boolean)
  On Error GoTo Hell

  If vNewValue = True Then
    '  m_Report.PaperSize =
    '   m_Report.PaperOrientation = crDefaultPaperOrientation
  End If

  Exit Property
Hell:

  MessageBox "PROC_PaperSizePrint_ERROR" & vbCrLf & vNewValue & vbCrLf & Err.Description, vbCritical, "Report Warning"
  Err.Clear
End Property

Public Property Let ReportTitle(ByVal vNewValue As Variant)
  mTitle = vNewValue
End Property

Private Sub cboPaper_Click(Index As Integer)

  If mLoad = True Then Exit Sub

End Sub

Private Sub cmdOk_Click(Index As Integer)

  'On Error Resume Next
  Select Case Index

    Case 0


      If Text1.Text = "" Then
        CRviewReport.SearchForText ("")
        MessageBox "Search Text not specified", "Search Text", msgOkOnly
      Else
        CRviewReport.SearchForText (Text1.Text)
      End If

    Case 1

      Picture1.Enabled = False
      Picture1.Visible = False

    Case 2

      Picture2.Enabled = False
      Picture2.Visible = False

    Case 3

      m_Report.PaperSize = cboPaper(0).ListIndex
      m_Report.PaperOrientation = cboPaper(1).ListIndex
      m_Report.PaperSource = cboPaper(2).ListIndex
      m_Report.TopMargin = IIf(TxtMargin(0) = "", 0.13, TxtMargin(0))
      m_Report.BottomMargin = IIf(TxtMargin(1) = "", 0.13, TxtMargin(1))
      m_Report.LeftMargin = IIf(TxtMargin(2) = "", 0.31, TxtMargin(2))
      m_Report.RightMargin = IIf(TxtMargin(3) = "", 0.13, TxtMargin(3))
      CRviewReport.Refresh

    Case 4

      m_Report.SelectPrinter m_Report.DriverName, cboPaper(2).Text, m_Report.PortName
      m_Report.PrintOut False, 1, False
  End Select

  Err.Clear
End Sub

Private Sub Form_Load()

  'On Error GoTo Hell
  If Not Rc Is Nothing Then
    If Rc.State = 1 Then
      If Rc.Recordcount <> 0 Then
        m_Report.DiscardSavedData
        m_Report.Database.SetDataSource Rc

        If mTitle <> "" Then m_Report.ReportTitle = mTitle
        m_Report.ReportComments = GetSetting(App.EXEName, "Lisence Profile", "Address") & vbCrLf & "Telp " & GetSetting( _
                App.EXEName, "Lisence Profile", "Phone") & vbCrLf & GetSetting(App.EXEName, "Lisence Profile", "City")

        If m_Report.ReportAuthor = "" Then m_Report.ReportAuthor = GetSetting(App.EXEName, "Lisence Profile", "Company Name")
        m_Report.EnableParameterPrompting = False
        Me.Caption = m_Report.ReportTitle
          
        CRviewReport.DisplayGroupTree = False
        CRviewReport.ReportSource = m_Report
        CRviewReport.ViewReport
        LblPages = migetCurrentPage
        KethokMas = MainMenu.SemeruTree1.Visible

        If KethokMas = True Then MainMenu.SemeruTree1.Visible = False
      End If
    End If
  End If

  CRviewReport.EnablePopupMenu = False
  'CRviewReport.EnableToolbar = False
Hell:

  'messagebox Err.Number & " - " & Err.Description
  Err.Clear
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

  If Not Rc Is Nothing Then
    If Rc.State = 1 Then Rc.Close
  End If

  Set Rc = Nothing
  Set m_Application = Nothing
  Set m_Report = Nothing

  If KethokMas = True Then MainMenu.SemeruTree1.Visible = True
  'Me.WindowState = vbMinimized
End Sub

Private Sub Form_Resize()

On Error Resume Next
  CRviewReport.Left = 0
  CRviewReport.Top = 0
  CRviewReport.width = Me.ScaleWidth
  CRviewReport.Height = Me.ScaleHeight
  Err.Clear
End Sub

Private Sub Picture1_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               Y As Single)
  MoveForm Picture1.hwnd
End Sub

Private Sub Picture2_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               Y As Single)
  MoveForm Picture2.hwnd
End Sub

Private Sub ShowPaperSize()
  Dim I As Integer
  Dim iCr As Boolean
  mLoad = True
  cboPaper(0).Clear
  CreateIdxPPR "Default", crDefaultPaperSize
  CreateIdxPPR "Letter", crPaperLetter
  CreateIdxPPR "Small Letter", crPaperLetterSmall
  CreateIdxPPR "Legal", crPaperLegal
  CreateIdxPPR "10x14", crPaper10x14
  CreateIdxPPR "11x17", crPaper11x17
  CreateIdxPPR "A3", crPaperA3
  CreateIdxPPR "A4", crPaperA4
  CreateIdxPPR "A4 Small", crPaperA4Small
  CreateIdxPPR "A5", crPaperA5
  CreateIdxPPR "B4", crPaperB4
  CreateIdxPPR "B5", crPaperB5
  CreateIdxPPR "C Sheet", crPaperCsheet
  CreateIdxPPR "D Sheet", crPaperDsheet
  CreateIdxPPR "Envelope 9", crPaperEnvelope9
  CreateIdxPPR "Envelope 10", crPaperEnvelope10
  CreateIdxPPR "Envelope 11", crPaperEnvelope11
  CreateIdxPPR "Envelope 12", crPaperEnvelope12
  CreateIdxPPR "Envelope 14", crPaperEnvelope14
  CreateIdxPPR "Envelope B4", crPaperEnvelopeB4
  CreateIdxPPR "Envelope B5", crPaperEnvelopeB5
  CreateIdxPPR "Envelope B6", crPaperEnvelopeB6
  CreateIdxPPR "Envelope C3", crPaperEnvelopeC3
  CreateIdxPPR "Envelope C4", crPaperEnvelopeC4
  CreateIdxPPR "Envelope C5", crPaperEnvelopeC5
  CreateIdxPPR "Envelope C6", crPaperEnvelopeC6
  CreateIdxPPR "Envelope C65", crPaperEnvelopeC65
  CreateIdxPPR "Envelope DL", crPaperEnvelopeDL
  CreateIdxPPR "Envelope Italy", crPaperEnvelopeItaly
  CreateIdxPPR "Envelope Monarch", crPaperEnvelopeMonarch
  CreateIdxPPR "Envelope Personal", crPaperEnvelopePersonal
  CreateIdxPPR "E Sheet", crPaperEsheet
  CreateIdxPPR "Executive", crPaperExecutive
  CreateIdxPPR "Fanfold Legal German", crPaperFanfoldLegalGerman
  CreateIdxPPR "Fanfold Standard German", crPaperFanfoldStdGerman
  CreateIdxPPR "Fanfold US", crPaperFanfoldUS
  CreateIdxPPR "Folio", crPaperFolio
  CreateIdxPPR "Ledger", crPaperLedger
  CreateIdxPPR "Note", crPaperNote
  CreateIdxPPR "Quarto", crPaperQuarto
  CreateIdxPPR "Statement", crPaperStatement
  CreateIdxPPR "Tabloid", crPaperTabloid

  With cboPaper(0)

    For I = 0 To .ListCount - 1

      If .ItemData(I) = m_Report.PaperSize Then
        .ListIndex = I
        iCr = True
      End If

    Next I

  End With

  If iCr = False Then
    CreateIdxPPR "Unknown Paper", m_Report.PaperSize
    cboPaper(0).Text = "Unknown Paper"
  End If

  cboPaper(1).Clear
  cboPaper(1).AddItem "Default Paper Orientation"
  cboPaper(1).AddItem "Portrait"
  cboPaper(1).AddItem "Landscape"
  cboPaper(1).ListIndex = m_Report.PaperOrientation
    
  ShowPrinterSource
  mLoad = False
  Picture2.Visible = True
  Picture2.Enabled = True
  'm_Report
  TxtMargin(0) = m_Report.TopMargin
  TxtMargin(1) = m_Report.BottomMargin
  TxtMargin(2) = m_Report.LeftMargin
  TxtMargin(3) = m_Report.RightMargin
End Sub

Private Sub CreateIdxPPR(Name As String, _
                         Index As Integer)
  cboPaper(0).AddItem Name
  cboPaper(0).ItemData(cboPaper(0).NewIndex) = Index
End Sub

Private Sub ShowPrinterSource()
  Dim I As Integer
  Dim PaperSource As Integer
  EnumPrinterBins m_Report.PrinterName, cboPaper(2)

  If cboPaper(3).ListIndex > 0 Then cboPaper(3).Text = m_Report.PrinterName

  With cboPaper(2)
    PaperSource = m_Report.PaperSource

    For I = 0 To .ListCount - 1

      If .ItemData(I) = PaperSource Then .ListIndex = I
    Next I

  End With

End Sub

Private Sub EnumPrinterBins(PrinterName As String, _
                            cbo As ComboBox)
  Dim prn As Printer
  Dim hPrinter As Long
  Dim dwbins As Long
  Dim I As Long
  Dim nameslist As String
  Dim NameBin As String
  Dim numBin() As Integer
  cboPaper(3).Clear
  cbo.Clear

  For Each prn In Printers
    cboPaper(3).AddItem prn.DeviceName

    If prn.DeviceName = PrinterName Then
      If OpenPrinter(prn.DeviceName, hPrinter, 0) <> 0 Then
        dwbins = DeviceCapabilities(prn.DeviceName, prn.Port, DC_BINS, ByVal vbNullString, 0)
        ReDim numBin(1 To dwbins)
        nameslist = String(24 * dwbins, 0)
        dwbins = DeviceCapabilities(prn.DeviceName, prn.Port, DC_BINS, numBin(1), 0)
        dwbins = DeviceCapabilities(prn.DeviceName, prn.Port, DC_BINNAMES, ByVal nameslist, 0)

        For I = 1 To dwbins
          NameBin = Mid(nameslist, 24 * (I - 1) + 1, 24)
          NameBin = Left(NameBin, InStr(1, NameBin, Chr(0)) - 1)
          cbo.AddItem NameBin
          cbo.ItemData(cbo.NewIndex) = numBin(I)
        Next I

        Call ClosePrinter(hPrinter)
      Else
        cbo.AddItem prn.DeviceName & "  <Unavailable>"
      End If
    End If

  Next prn

End Sub

Private Function migetCurrentPage() As Integer
  On Error GoTo Hell
  While CRviewReport.IsBusy

    DoEvents
  Wend
  migetCurrentPage = CRviewReport.GetCurrentPageNumber
Hell:

  Err.Clear
End Function

Private Sub TutupReport()
  Dim frm As Form
  On Error GoTo Hell

  For Each frm In Forms

    If UCase(frm.Tag) = "RPT" Then
      Unload frm
    End If

  Next

  Set frm = Nothing
Hell:

  Err.Clear
End Sub

