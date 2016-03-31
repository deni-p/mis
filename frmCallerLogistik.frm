VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCallerLogistik 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11325
   Icon            =   "frmCallerLogistik.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   Tag             =   "Data Viewer"
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   11265
      TabIndex        =   7
      Top             =   7575
      Width           =   11325
      Begin VB.CommandButton cmdSet 
         Appearance      =   0  'Flat
         Caption         =   "Kirim Ke SPP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   7425
         Picture         =   "frmCallerLogistik.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   90
         Width           =   1365
      End
      Begin VB.TextBox txtCari 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   75
         TabIndex        =   2
         Top             =   285
         Width           =   3735
      End
      Begin VB.CommandButton CmdOK 
         Appearance      =   0  'Flat
         Caption         =   "P&ilih"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   9630
         Picture         =   "frmCallerLogistik.frx":6BDC
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   75
         Width           =   800
      End
      Begin VB.CommandButton CmdLink 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   4680
         Picture         =   "frmCallerLogistik.frx":7166
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   135
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton CmdExit 
         Appearance      =   0  'Flat
         Caption         =   "&Keluar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   10425
         Picture         =   "frmCallerLogistik.frx":74F0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   75
         Width           =   800
      End
      Begin VB.CommandButton cmdRefresh 
         Appearance      =   0  'Flat
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   8835
         Picture         =   "frmCallerLogistik.frx":7A7A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   75
         Width           =   800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cari Kriteria"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   75
         TabIndex        =   1
         Top             =   45
         Width           =   840
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7530
      Left            =   0
      ScaleHeight     =   7530
      ScaleWidth      =   11325
      TabIndex        =   8
      Top             =   0
      Width           =   11325
      Begin VB.Frame FrTombol 
         Height          =   30
         Left            =   45
         TabIndex        =   10
         Top             =   7425
         Width           =   11220
      End
      Begin VB.PictureBox Picture1 
         Height          =   7335
         Left            =   15
         ScaleHeight     =   7275
         ScaleWidth      =   11205
         TabIndex        =   9
         Top             =   15
         Width           =   11265
         Begin MSDataGridLib.DataGrid DgDetail 
            Height          =   7245
            Left            =   0
            TabIndex        =   0
            Top             =   0
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   12779
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BorderStyle     =   0
            HeadLines       =   2
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
               DataField       =   "tanggal"
               Caption         =   "Tanggal"
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
               DataField       =   "InternalName"
               Caption         =   "Barang"
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
               DataField       =   "Quote Qty"
               Caption         =   "Qty Diminta"
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
            BeginProperty Column03 
               DataField       =   "UOM"
               Caption         =   "Satuan"
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
            BeginProperty Column04 
               DataField       =   "Issued By"
               Caption         =   "Oleh"
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
            BeginProperty Column05 
               DataField       =   "penggunaan"
               Caption         =   "Penggunaan"
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
               DataField       =   "description"
               Caption         =   "Tgl Kebutuhan"
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
               MarqueeStyle    =   3
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
               BeginProperty Column02 
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column04 
               EndProperty
               BeginProperty Column05 
               EndProperty
               BeginProperty Column06 
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSMask.MaskEdBox MaskCari 
      Height          =   270
      Left            =   4950
      TabIndex        =   11
      Top             =   5040
      Visible         =   0   'False
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "frmCallerLogistik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public IsYes As Integer
Private WithEvents mVarFormObject As Recordset
Attribute mVarFormObject.VB_VarHelpID = -1
Private mLoop As Integer
Private mvarTagForm As String
Private myStd As New StdDataFormat
Private mVarBoolean, mVarLoad As Boolean
Private MyFrm As Form
Private LookupCaller As Form
Event RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Event ActionJackson(ByVal ColIndex As Integer)
Event CallLinkForm()
Event BeforeUnload()
Dim yRow, xCol As Integer

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdLink_Click()
RaiseEvent CallLinkForm
End Sub

Private Sub cmdRefresh_Click()
On Error Resume Next
Set mVarFormObject.ActiveConnection = CNN
mVarFormObject.Filter = adFilterNone
mVarFormObject.Requery
'Tampil
txtCari.Text = ""
txtCari.SetFocus
Err.Clear
End Sub

Private Sub Command1_Click()
Call dgDetail_DblClick
End Sub

Private Sub cmdSet_Click()
    If MessageBox("Apakah Benar akan diajukan ke SPP ?", "Konfirmasi", msgYesNo, msgQuestion) = 1 Then
      Dim nIDGen As New IDGenerator
      Dim xID As String
      xID = nIDGen.GetID("PP")
      SendDataToServer " INSERT INTO  SPP_Header ( SPPID," & _
                                       "SPP_Date," & _
                                       "Note," & _
                                       "Ordered_by," & _
                                       "Status)" & _
            " Values ('" & xID & _
                   "','" & Format(Now, "yyyy-MM-dd") & _
                   "','" & "-" & _
                   "','" & MainMenu.StatusBar1.Panels(1).Text & _
                   "', 0)"
      
       SendDataToServer " INSERT INTO SPP_Line ( SPPID, NoItem, QTY_SPP, Keperluan, Note) " & _
                               " VALUES (N'" & xID & "', N'" & mVarFormObject.Fields("Item ID") & "', " & FQty(mVarFormObject.Fields("Quote Qty")) & ", N'-','-')"
      
      MessageBox "Permintaan Pembelian telah dibuat", "Konfirmasi", msgOkOnly, msgInfo
      Set nIDGen = Nothing
    End If
End Sub

Private Sub DGDETAIL_ButtonClick(ByVal ColIndex As Integer)
   RaiseEvent ActionJackson(ColIndex)
End Sub

Private Sub DgDetail_Click()
On Error Resume Next
mLoop = xCol
Label1 = "&Cari kriteria  berdasarkan " & UCase(dgDetail.Columns(mLoop).DataField)
dgDetail.col = xCol
Err.Clear
End Sub

Private Sub dgDetail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   RaiseEvent RowColChange(mvarTagForm, mVarFormObject)
   mVarBoolean = True
End If
End Sub

Private Sub DgDetail_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
xCol = dgDetail.ColContaining(x)
yRow = dgDetail.RowContaining(Y)
End Sub

Private Sub Form_Activate()
Me.Caption = mvarTagForm
Me.Tag = mvarTagForm
HiasFormManTell Picture2, Me
mVarLoad = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
       Case vbKeyReturn:
           If mVarBoolean = True Then
              Unload Me
           Else
              mVarBoolean = False
           End If
End Select
End Sub

Private Sub Form_Paint()
On Error Resume Next
If mVarLoad = True Then Exit Sub
HiasFormManTell Picture2, Me
'HiasFormCaller Picture1, Me
Me.Caption = ""
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmCaller = Nothing
End Sub

Private Sub Label1_Click()
txtCari.SetFocus
End Sub

Private Sub mVarFormObject_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'RaiseEvent RowColChange(mvarTagForm, pRecordset)
End Sub

Public Property Get FromTagActive() As String
       FromTagActive = mvarTagForm
       'FromTagActive = Me.Tag
End Property

Public Property Let FromTagActive(ByVal vNewValue As String)
       mvarTagForm = vNewValue
       Me.Tag = mvarTagForm
       Me.Caption = mvarTagForm
End Property

Private Sub cmdOk_Click()
'RaiseEvent BeforeUnload
dgDetail_DblClick
Unload Me
End Sub

Private Sub dgDetail_DblClick()
   RaiseEvent RowColChange(mvarTagForm, mVarFormObject)
   Unload Me
End Sub

Private Sub dgDetail_Error(ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub

Private Sub dgDetail_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
If ColIndex <= 0 Then
   mLoop = 0
Else
   mLoop = ColIndex
End If
Label1 = "&Cari kriteria  berdasarkan " & UCase(dgDetail.Columns(ColIndex).DataField)
dgDetail.col = ColIndex
Err.Clear
txtCari.SetFocus
End Sub


Private Sub Form_Load()
CallerID = True
mLoop = 0
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
'mvarTagForm = Me.Tag
'Me.Caption = mvarTagForm

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If mVarBoolean = True Then
'
'End If
CloseDB mVarFormObject
Set mVarFormObject = Nothing
Set myStd = Nothing
CallerID = False
If Not LookupCaller Is Nothing Then
    LookupCaller.Enabled = True
    LookupCaller.ZOrder
End If
Set LookupCaller = Nothing
RaiseEvent BeforeUnload
End Sub

Private Sub Form_Resize()
'On Error Resume Next
HiasFormManTell Picture2, Me
'HiasForm Picture1, Me
'Me.Caption = ""
''Me.Height = 6855
''Me.Width = 7470
'Err.Clear
End Sub

Public Property Let CaptionLink(ByVal AliasForm As String)
'If Not MyFrm Is Nothing Then
   If AliasForm <> "" Then
    cmdLink.Caption = AliasForm
    cmdLink.Enabled = True
    cmdLink.Visible = True
   End If
'End If
End Property

Public Property Set FormData(ByVal vData As Recordset)
    Me.Caption = mvarTagForm
    Set mVarFormObject = vData
    Set dgDetail.DataSource = mVarFormObject
 '   AdjustDataGridColumns DgDetail, mVarFormObject, mVarFormObject.Recordcount, mVarFormObject.Fields.Count, True
    'Tampil
End Property

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

'Private Sub txtCari_Change()
'Dim StrCari As String
'If Not mVarFormObject Is Nothing Then
'   If mVarFormObject.Recordcount <> 0 Then
'      If txtCari <> "" Then
'         StrCari = "[" & DgDetail.Columns(mLoop).DataField & "]" & " Like '" & txtCari & "%'"
'         mVarFormObject.Find StrCari, 0, adSearchForward, adBookmarkFirst
'         If mVarFormObject.EOF Then
'            messagebox "Kriteria Yang Dicari Tidak Ada..............!", vbCritical
'         End If
'      End If
'   End If
'End If
'End Sub

Private Sub TxtCari_Change()
Dim strcari As String
If txtCari <> "" Then
    If Not mVarFormObject Is Nothing Then
       If mVarFormObject.Recordcount <> 0 Then
         Select Case mVarFormObject.Fields(dgDetail.Columns(mLoop).DataField).Type
            Case adBigInt, adInteger, adCurrency, adDecimal, adDouble, adNumeric, adSingle, adSmallInt, adTinyInt, adVarNumeric
               strcari = "[" & dgDetail.Columns(mLoop).DataField & "]" & " = " & txtCari
            Case Else
               strcari = "[" & dgDetail.Columns(mLoop).DataField & "]" & " Like '" & txtCari & "%'"
         End Select
          mVarFormObject.Filter = strcari ', 0, adSearchForward, adBookmarkFirst
          If mVarFormObject.Recordcount = 0 Then MessageBox "Kriteria Yang Dicari Tidak Ada..............!", vbCritical
       Else
          Call cmdRefresh_Click
       End If
    End If
Else
    Call cmdRefresh_Click
End If
End Sub
Private Sub txtCari_KeyPress(KeyAscii As Integer)
If mVarFormObject.Fields(dgDetail.Columns(mLoop).DataField).Type = 3 Then
    ValidNum KeyAscii
End If
End Sub
Public Property Let FindCaller(ByVal FindStringValue As String)
On Error GoTo Hell
If Not mVarFormObject Is Nothing Then
   If mVarFormObject.State = 1 Then
      If FindStringValue <> "" Then
         mVarFormObject.Find FindStringValue, 0, adSearchForward, adBookmarkFirst
         If mVarFormObject.EOF Then
            mVarFormObject.AbsolutePosition = 1
         End If
      End If
   End If
End If
Hell:
    Err.Clear
End Property

Public Property Get GetFieldByName(ByVal FldListFields As Variant) As Variant
On Error GoTo Hell
If mVarFormObject.Recordcount <> 0 Then
   GetFieldByName = IIf(Not IsNull(mVarFormObject.Fields(FldListFields).Value), mVarFormObject.Fields(FldListFields).Value, Empty)
Else
   GetFieldByName = Empty
End If
Exit Sub
Hell:
Err.Clear
End Property

Public Property Let CreateButton(ByVal ColIndex As Integer, ByVal vNewValue As Boolean)
'DgDetail.Columns(ColIndex).Button = vNewValue
End Property

Public Property Let SetFormat(ByVal ColIndex As Integer, ByVal vNewValue As String)
'myStd.Type = fmtCustom
'myStd.Format = vNewValue
'Set DgDetail.Columns(ColIndex).DataFormat = myStd
End Property

Public Property Let SetAlignmentFormat(ByVal ColIndex As Integer, ByVal vNewValue As Integer)
'DgDetail.Columns(ColIndex).Alignment = vNewValue
End Property

Private Sub OpenKolom()
Dim I As Integer
Dim Avdata As Variant
I = 0
'If mVarFormObject.Recordcount <> 0 Then
'  mVarFormObject.MoveFirst
'  DgDetail.Cols = mVarFormObject.Fields.Count + 1
'  DgDetail.Rows = mVarFormObject.Recordcount + 1
'  Avdata = mVarFormObject.GetString(adClipString, mVarFormObject.Recordcount)
'
'  For I = 0 To DgDetail.Cols - 2
'      DgDetail.TextMatrix(0, I + 1) = mVarFormObject.Fields(I).Name
'  Next
'  DgDetail.Row = 1
'  DgDetail.Col = 1
'  DgDetail.RowSel = DgDetail.Rows - 1
'  DgDetail.ColSel = DgDetail.Cols - 1
'  DgDetail.Clip = Avdata
'  mVarFormObject.MoveFirst
'  DgDetail.RowSel = DgDetail.Row
'  DgDetail.ColSel = DgDetail.Col
'  DoColumnSort
'End If
End Sub

'Sub DoColumnSort()
'    On Error Resume Next
'    With DgDetail
'        .Redraw = False
'        .Row = 1
'        .RowSel = .Rows - 1
'        .FillStyle = flexFillRepeat
'        .Col = .FixedCols + 1
'        .Row = .FixedRows
'        .RowSel = .Rows - 1
'        .ColSel = .Cols '- 1
'        .CellBackColor = &HFCF1ED
'        Dim iLoop As Integer
'        For iLoop = .FixedRows + 1 To .Rows - 1 Step 2
'            .Row = iLoop
'            .Col = .FixedCols
'            .ColSel = .Cols() - .FixedCols
'            .CellBackColor = &HEAAF6F
'            .GridColor = &H6D4016
'        Next iLoop
'        .FillStyle = flexFillSingle
'        .Redraw = True
'    End With
'
'End Sub

'Private Sub DrawIcons(ByVal PicType As PictureType, ByVal vRow As Long, Optional UseRedraw As Boolean = True)
'    On Error GoTo Err_DrawIcons
'    Dim Irow&, iCol&, iRowsel&, iColSel&
'    With MSHFlexGrid1
'        If UseRedraw Then .Redraw = False
'        Irow = .Row
'        iCol = .Col
'        iRowsel = .RowSel
'        iColSel = .ColSel
'        .Row = vRow
'        .Col = 0
'        .CellPictureAlignment = flexAlignCenterCenter
'        If PicType = ptNone Then
'            Set .CellPicture = Nothing
'        Else
'            Set .CellPicture = imgHolders(PicType).Picture
'        End If
'        .Row = Irow
'        .Col = iCol
'        .RowSel = iRowsel
'        .ColSel = iColSel
'        If UseRedraw Then .Redraw = True
'    End With
'    Exit Sub
'Err_DrawIcons:
'    ErrorMsg Err.Number, Err.Description, "DrawIcons", mcstrMod
'End Sub

Private Sub Tampil()
On Error GoTo 3
    Dim I As Integer
    Label1 = "&Cari Kriteria by " & mVarFormObject.Fields(I).Name
    'DgDetail.Height = (DgDetail.RowHeight * 19) + 50
    For I = 0 To dgDetail.Columns.Count - 1
        If (dgDetail.Columns.Count - 1) >= 1 Then
            If (dgDetail.RowHeight * mVarFormObject.Recordcount) > dgDetail.Height Then
                If (dgDetail.Columns.Count - 1) > 2 Then
                   dgDetail.Columns(I).width = ((dgDetail.width / 1.8) / 2) + 184
                ElseIf (dgDetail.Columns.Count - 1) <= 2 Then
                   dgDetail.Columns(I).width = (dgDetail.width / 2) - 280
                End If
            Else
                If (dgDetail.Columns.Count) = 2 Then
                    dgDetail.Columns(I).width = (dgDetail.width / 2) - 300
                ElseIf (dgDetail.Columns.Count - 1) >= 2 Then
                   dgDetail.Columns(I).width = ((dgDetail.width / 1.8) / 2) + 250
                ElseIf (dgDetail.Columns.Count - 1) <= 2 Then
                   dgDetail.Columns(I).width = (dgDetail.width / 2) - 190
                End If
            End If
        Else
           dgDetail.Columns(I).width = dgDetail.width - 600
        End If
        dgDetail.Columns(I).DividerStyle = dbgLightGrayLine
        dgDetail.Columns(I).Locked = True
        With mVarFormObject
            Select Case .Fields(I).Type
                   Case adBigInt, adCurrency, adDecimal, adDouble, adInteger
                        myStd.Type = fmtCustom
                        myStd.Format = "#,##0"
                        dgDetail.Columns(I).NumberFormat = "#,##0"
                        dgDetail.Columns(I).Alignment = dbgRight
                   Case adDate, adDBDate, adDBTime, adDBTimeStamp
'                        Debug.Print .Fields(I).Value
'                        myStd.Type = fmtCustom
'                        myStd.Format = "dd mm yyyy"
'                        Set DGDetail.Columns(I).DataFormat = myStd
                        dgDetail.Columns(I).Alignment = dbgLeft
                   Case 11:
                        myStd.Type = fmtBoolean
                        myStd.Format = "YES/NO"
                        Set dgDetail.Columns(I).DataFormat = myStd
                        dgDetail.Columns(I).Alignment = dbgRight
                   Case Else:
            End Select
        End With
    Next I
    dgDetail.ToolTipText = Me.Caption
Exit Sub
3:
MessageBox Err.Description, "frmcaller:tampil" & Err.Number, msgOkOnly, msgExclamation
End Sub

Public Function LookUp(Caller As Form)
    On Error GoTo 2
    If LookupCaller Is Nothing Then
        Caller.Enabled = False
        Set LookupCaller = Caller
    Else
        MessageBox Me.Caption & " already in use." & vbCrLf & "Please complete prvious request."
    End If
    Me.Show
    Me.ZOrder
Exit Function
2:
MessageBox Err.Description, "frmcaller:lookup" & Err.Number, msgOkOnly, msgExclamation

End Function

Public Sub AdjustDataGridColumns _
           (DG As DataGrid, _
           adoData As Recordset, _
           intRecord As Integer, _
           intField As Integer, _
           Optional AccForHeaders As Boolean)

'This procedure will adjust DataGrids column width
'based on longest field in underlying source

'DG = DataGrid
'adoData = Adodc control
'intRecord = Number of record
'intField = Number of field
'AccForHeaders = True or False

    Dim row As Long, col As Long
    Dim width As Single, maxWidth As Single
    Dim saveFont As StdFont, saveScaleMode As Integer
    Dim cellText As String
    
    'If number of records = 0 then exit from the sub
    If intRecord = 0 Then Exit Sub
    'Save the form's font for DataGrid's font
    'We need this for form's TextWidth method
    Set saveFont = DG.Parent.Font
    Set DG.Parent.Font = DG.Font
    'Adjust ScaleMode to vbTwips for the form (parent).
    saveScaleMode = DG.Parent.ScaleMode
    DG.Parent.ScaleMode = vbTwips
    'Always from first record...
    adoData.MoveFirst
    maxWidth = 0
    'We begin from the first column until the last column
    For col = 0 To intField - 1
        adoData.MoveFirst
        'Optional param, if true, set maxWidth to
        'width of DG.Parent
        If AccForHeaders Then
            maxWidth = DG.Parent.TextWidth(DG.Columns(col).Text) + 200
        End If
        'Repeat from first record again after we have
        'finished process the last record in
        'former column...
        adoData.MoveFirst
        For row = 0 To intRecord - 1
            'Get the text from the DataGrid's cell
            If intField = 1 Then
            Else  'If number of field more than one
                cellText = DG.Columns(col).Text
            End If
            'Fix the border...
            'Not for "multiple-line text"...
            width = DG.Parent.TextWidth(cellText) + 200
            'Update the maximum width if we found
            'the wider string...
            If width > maxWidth Then
               maxWidth = width
               DG.Columns(col).width = maxWidth
            End If
            'Process next record...
            adoData.MoveNext
        Next row
        'Change the column width...
        DG.Columns(col).width = maxWidth 'kolom terakhir!
    Next col
    'Change the DataGrid's parent property
    Set DG.Parent.Font = saveFont
    DG.Parent.ScaleMode = saveScaleMode
    'If finished, then move pointer to first record again
    adoData.MoveFirst
End Sub  'End of AdjustDataGridColumns






