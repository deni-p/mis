VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form fCrwViewer 
   Caption         =   "Viewer"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9210
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fCrwViewer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin ComCtl3.CoolBar CBar 
      Height          =   420
      Left            =   765
      TabIndex        =   6
      Top             =   7035
      Visible         =   0   'False
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   741
      BandCount       =   1
      BandBorders     =   0   'False
      VariantHeight   =   0   'False
      _CBWidth        =   6000
      _CBHeight       =   420
      _Version        =   "6.0.8169"
      MinHeight1      =   360
      Width1          =   2880
      NewRow1         =   0   'False
   End
   Begin VB.Frame FrTools 
      Caption         =   " Tools "
      Height          =   5235
      Left            =   45
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   810
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   4020
         Left            =   150
         TabIndex        =   2
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
      Begin VB.Label LblPages 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0/0"
         Height          =   195
         Left            =   255
         TabIndex        =   5
         Top             =   4800
         Width           =   240
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Page "
         Height          =   195
         Left            =   195
         TabIndex        =   4
         Top             =   4560
         Width           =   420
      End
      Begin VB.Label LblCap 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   60
         TabIndex        =   3
         Top             =   7440
         Width           =   1125
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRviewReport 
      Height          =   6810
      Left            =   1125
      TabIndex        =   0
      Top             =   75
      Width           =   7920
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   -1  'True
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   105
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCrwViewer.frx":6852
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCrwViewer.frx":D0B4
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCrwViewer.frx":13916
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCrwViewer.frx":1A178
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCrwViewer.frx":209DA
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCrwViewer.frx":2723C
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCrwViewer.frx":2DA9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCrwViewer.frx":34300
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCrwViewer.frx":3AB62
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCrwViewer.frx":413C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCrwViewer.frx":47C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCrwViewer.frx":4E488
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCrwViewer.frx":54CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCrwViewer.frx":5B54C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCrwViewer.frx":61DAE
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "fCrwViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rc As New Recordset
Private m_Application As New CRAXDDRT.Application
Private mStr As String


Private Sub Form_Activate()
On Error GoTo Hell
If Rc Is Nothing Then GoTo Hell
If Rc.RecordCount <= 0 Then
   MsgBox "Can't View Report please try another period.", vbInformation, "No Record Detected"
   Unload Me
End If

Exit Sub
Hell:
    MsgBox "Can't View Report please try another period.", vbInformation, "No Record Detected"
    Unload Me
    Err.Clear
End Sub

Private Sub Form_Load()
Dim MyCtrl As Object
On Error Resume Next
If Not Rc Is Nothing Then
   If Rc.State = 1 Then If Rc.RecordCount <> 0 Then RefreshReport
End If
CRviewReport.EnablePopupMenu = False
  
'' Add a new CommandButton to the first Band.
'Set MyCtrl = Controls.Add("VB.CommandButton", "cmdTest", CBar)
'MyCtrl.Caption = "Printer Setup"
'Set CBar.Bands(1).Child = MyCtrl  ' place on first Band
'' Add a new TextBox to the second Band.
'Set MyCtrl = Controls.Add("VB.TextBox", "txtTest", CBar)
'MyCtrl.Text = "Testing Text"
'Set CBar.Bands(2).Child = MyCtrl  ' place on second Band

Err.Clear
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Set Rc = Nothing
Set m_Report = Nothing
Set m_Application = Nothing

Err.Clear
End Sub

Private Sub Form_Resize()
On Error Resume Next
CRviewReport.Left = 0
CRviewReport.Top = 0
CRviewReport.Width = Me.ScaleWidth
CRviewReport.Height = Me.ScaleHeight
Err.Clear
'
'On Error Resume Next
'FrTools.Left = 20
'FrTools.Top = 0
'CRviewReport.Top = 0
'CRviewReport.Left = FrTools.Width + 50
'CRviewReport.Height = Me.ScaleHeight
'CRviewReport.Width = Me.ScaleWidth - (FrTools.Width + 50)
'FrTools.Height = Me.ScaleHeight
'FrTools.Left = 20
'FrTools.Top = 0
'
'CRviewReport.Top = 0
'CRviewReport.FrTools.Width 50   'Left = 0 '
'CRviewReport.Height = Me.ScaleHeight
'CRviewReport.Width = Me.ScaleWidth - 50
'FrTools.Height = Me.ScaleHeight
'Err.Clear
End Sub

Public Sub CallReport(ByVal SelectQuery As String, ByVal ReportName As String, ByVal ReportFilePath As String, ByVal ReportTitle As String)
Dim Myrpt As New Utility
mStr = SelectQuery
Set Rc = Myrpt.OpenDB(SelectQuery)
With Rc
     Set m_Report = m_Application.OpenReport(ReportFilePath & "\" & ReportName)
     If ReportTitle <> "" Then m_Report.ReportTitle = ReportTitle

End With
Set Myrpt = Nothing
End Sub

Private Sub RefreshReport()
m_Report.DiscardSavedData
m_Report.Database.SetDataSource Rc
m_Report.EnableParameterPrompting = False
m_Report.ReadRecords
Me.Caption = m_Report.ReportTitle
CRviewReport.DisplayGroupTree = False
CRviewReport.ReportSource = m_Report
CRviewReport.ViewReport
End Sub

Private Sub ReloadReport()
Dim Myrpt As New Utility
Set Rc = Myrpt.OpenDB(mStr)
RefreshReport
Set Myrpt = Nothing
End Sub

Private Sub CRviewReport_RefreshButtonClicked(UseDefault As Boolean)
If Not CRviewReport.IsBusy Then ReloadReport
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmClosing.Toolbar1.Buttons(14).Visible = False
End Sub
'
'Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error Resume Next
'Select Case Button.Index
'       Case 1:
'          If CRviewReport.ActiveViewIndex > 1 Then
'             CRviewReport.CloseView (CRviewReport.ActiveViewIndex)
'          End If
'       Case 2: CRviewReport.PrintReport
'       Case 3: m_Report.Export
'       Case 4: CRviewReport.Refresh
'       Case 5: CRviewReport.DisplayGroupTree = Not CRviewReport.DisplayGroupTree
'       Case 6: CRviewReport.Zoom 100
'       Case 7:
'            Picture1.Enabled = True
'            Picture1.Visible = True
'       Case 8:
'            CRviewReport.ShowFirstPage
'            LblPages = migetCurrentPage
'       Case 9:
'            CRviewReport.ShowNextPage
'            LblPages = migetCurrentPage
'       Case 10:
'            CRviewReport.ShowPreviousPage
'            LblPages = migetCurrentPage
'       Case 11:
'            CRviewReport.ShowLastPage
'            LblPages = migetCurrentPage
'       Case 12: Unload Me
'End Select
'Err.Clear
'End Sub
'
'Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
'Dim CboPPr As CRPaperSize
'Select Case ButtonMenu
'       Case "Paper Fit": CRviewReport.Zoom 100
'       Case "100%": CRviewReport.Zoom 100
'       Case "90%": CRviewReport.Zoom 90
'       Case "80%": CRviewReport.Zoom 80
'       Case "70%": CRviewReport.Zoom 70
'       Case "60%": CRviewReport.Zoom 60
'       Case "50%": CRviewReport.Zoom 50
'       Case "40%": CRviewReport.Zoom 40
'       Case "30%": CRviewReport.Zoom 30
'       Case "20%": CRviewReport.Zoom 20
'       Case "10%": CRviewReport.Zoom 10
'       Case "Printer Properties":
'            m_Report.PrinterSetup Me.hwnd
'            CRviewReport.Refresh
'       Case "Send Document To Printer": m_Report.PrintOut False, 1
'       Case "Papersize Option": ShowPaperSize
'End Select
'End Sub
'
'
'Private Sub ShowPaperSize()
'    Dim I As Integer
'    Dim iCr As Boolean
'    mLoad = True
'    cboPaper(0).Clear
'    CreateIdxPPR "Default", crDefaultPaperSize
'    CreateIdxPPR "Letter", crPaperLetter
'    CreateIdxPPR "Small Letter", crPaperLetterSmall
'    CreateIdxPPR "Legal", crPaperLegal
'    CreateIdxPPR "10x14", crPaper10x14
'    CreateIdxPPR "11x17", crPaper11x17
'    CreateIdxPPR "A3", crPaperA3
'    CreateIdxPPR "A4", crPaperA4
'    CreateIdxPPR "A4 Small", crPaperA4Small
'    CreateIdxPPR "A5", crPaperA5
'    CreateIdxPPR "B4", crPaperB4
'    CreateIdxPPR "B5", crPaperB5
'    CreateIdxPPR "C Sheet", crPaperCsheet
'    CreateIdxPPR "D Sheet", crPaperDsheet
'    CreateIdxPPR "Envelope 9", crPaperEnvelope9
'    CreateIdxPPR "Envelope 10", crPaperEnvelope10
'    CreateIdxPPR "Envelope 11", crPaperEnvelope11
'    CreateIdxPPR "Envelope 12", crPaperEnvelope12
'    CreateIdxPPR "Envelope 14", crPaperEnvelope14
'    CreateIdxPPR "Envelope B4", crPaperEnvelopeB4
'    CreateIdxPPR "Envelope B5", crPaperEnvelopeB5
'    CreateIdxPPR "Envelope B6", crPaperEnvelopeB6
'    CreateIdxPPR "Envelope C3", crPaperEnvelopeC3
'    CreateIdxPPR "Envelope C4", crPaperEnvelopeC4
'    CreateIdxPPR "Envelope C5", crPaperEnvelopeC5
'    CreateIdxPPR "Envelope C6", crPaperEnvelopeC6
'    CreateIdxPPR "Envelope C65", crPaperEnvelopeC65
'    CreateIdxPPR "Envelope DL", crPaperEnvelopeDL
'    CreateIdxPPR "Envelope Italy", crPaperEnvelopeItaly
'    CreateIdxPPR "Envelope Monarch", crPaperEnvelopeMonarch
'    CreateIdxPPR "Envelope Personal", crPaperEnvelopePersonal
'    CreateIdxPPR "E Sheet", crPaperEsheet
'    CreateIdxPPR "Executive", crPaperExecutive
'    CreateIdxPPR "Fanfold Legal German", crPaperFanfoldLegalGerman
'    CreateIdxPPR "Fanfold Standard German", crPaperFanfoldStdGerman
'    CreateIdxPPR "Fanfold US", crPaperFanfoldUS
'    CreateIdxPPR "Folio", crPaperFolio
'    CreateIdxPPR "Ledger", crPaperLedger
'    CreateIdxPPR "Note", crPaperNote
'    CreateIdxPPR "Quarto", crPaperQuarto
'    CreateIdxPPR "Statement", crPaperStatement
'    CreateIdxPPR "Tabloid", crPaperTabloid
'    With cboPaper(0)
'        For I = 0 To .ListCount - 1
'            If .ItemData(I) = m_Report.PaperSize Then
'               .ListIndex = I
'               iCr = True
'            End If
'        Next I
'    End With
'    If iCr = False Then
'       CreateIdxPPR "Unknown Paper", m_Report.PaperSize
'       cboPaper(0).Text = "Unknown Paper"
'    End If
'    cboPaper(1).Clear
'    cboPaper(1).AddItem "Default Paper Orientation"
'    cboPaper(1).AddItem "Portrait"
'    cboPaper(1).AddItem "Landscape"
'    cboPaper(1).ListIndex = m_Report.PaperOrientation
'
'    ShowPrinterSource
'    mLoad = False
'    Picture2.Visible = True
'    Picture2.Enabled = True
'    'm_Report
'    TxtMargin(0) = m_Report.TopMargin
'    TxtMargin(1) = m_Report.BottomMargin
'    TxtMargin(2) = m_Report.LeftMargin
'    TxtMargin(3) = m_Report.RightMargin
'End Sub
'
'Private Sub CreateIdxPPR(Name As String, Index As Integer)
'    cboPaper(0).AddItem Name
'    cboPaper(0).ItemData(cboPaper(0).NewIndex) = Index
'End Sub
'
'Private Sub ShowPrinterSource()
'    Dim I As Integer
'    Dim PaperSource As Integer
'    EnumPrinterBins m_Report.PrinterName, cboPaper(2)
'    If cboPaper(3).ListIndex > 0 Then cboPaper(3).Text = m_Report.PrinterName
'    With cboPaper(2)
'        PaperSource = m_Report.PaperSource
'        For I = 0 To .ListCount - 1
'            If .ItemData(I) = PaperSource Then .ListIndex = I
'        Next I
'    End With
'End Sub
'
'Private Sub EnumPrinterBins(PrinterName As String, cbo As ComboBox)
'    Dim prn As Printer
'    Dim hPrinter As Long
'    Dim dwbins As Long
'    Dim I As Long
'    Dim nameslist As String
'    Dim NameBin As String
'    Dim numBin() As Integer
'    cboPaper(3).Clear
'    cbo.Clear
'    For Each prn In Printers
'        cboPaper(3).AddItem prn.DeviceName
'        If prn.DeviceName = PrinterName Then
'            If OpenPrinter(prn.DeviceName, hPrinter, 0) <> 0 Then
'                dwbins = DeviceCapabilities(prn.DeviceName, prn.Port, DC_BINS, ByVal vbNullString, 0)
'                ReDim numBin(1 To dwbins)
'                nameslist = String(24 * dwbins, 0)
'                dwbins = DeviceCapabilities(prn.DeviceName, prn.Port, DC_BINS, numBin(1), 0)
'                dwbins = DeviceCapabilities(prn.DeviceName, prn.Port, DC_BINNAMES, ByVal nameslist, 0)
'                For I = 1 To dwbins
'                    NameBin = Mid(nameslist, 24 * (I - 1) + 1, 24)
'                    NameBin = Left(NameBin, InStr(1, NameBin, Chr(0)) - 1)
'                    cbo.AddItem NameBin
'                    cbo.ItemData(cbo.NewIndex) = numBin(I)
'                Next I
'                Call ClosePrinter(hPrinter)
'            Else
'                cbo.AddItem prn.DeviceName & "  <Unavailable>"
'            End If
'        End If
'    Next prn
'End Sub
'
'Private Function migetCurrentPage() As Integer
'On Error GoTo Hell
'    While CRviewReport.IsBusy
'        DoEvents
'    Wend
'    migetCurrentPage = CRviewReport.GetCurrentPageNumber
'Hell:
'    Err.Clear
'End Function
'
'Private Sub TutupReport()
'Dim Frm As Form
'On Error GoTo Hell
'For Each Frm In Forms
'    If UCase(Frm.Tag) = "RPT" Then
'       Unload Frm
'    End If
'Next
'Set Frm = Nothing
'Hell:
'    Err.Clear
'End Sub
