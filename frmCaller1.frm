VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmCaller1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8685
   Icon            =   "frmCaller1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   Tag             =   "Data Viewer"
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   8625
      TabIndex        =   6
      Top             =   5685
      Width           =   8685
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
         TabIndex        =   1
         Top             =   285
         Width           =   3000
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
         Left            =   7005
         Picture         =   "frmCaller1.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
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
         Left            =   4845
         Picture         =   "frmCaller1.frx":6DDC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   60
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
         Left            =   7800
         Picture         =   "frmCaller1.frx":7166
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
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
         Left            =   6210
         Picture         =   "frmCaller1.frx":76F0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
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
         TabIndex        =   0
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
      Height          =   5670
      Left            =   0
      ScaleHeight     =   5670
      ScaleWidth      =   8685
      TabIndex        =   7
      Top             =   0
      Width           =   8685
      Begin VB.Frame FrTombol 
         Height          =   30
         Left            =   45
         TabIndex        =   9
         Top             =   5565
         Width           =   8580
      End
      Begin VB.PictureBox Picture1 
         Height          =   5490
         Left            =   15
         ScaleHeight     =   5430
         ScaleWidth      =   8580
         TabIndex        =   8
         Top             =   15
         Width           =   8640
         Begin TrueOleDBGrid80.TDBGrid dgDetail 
            Height          =   5385
            Left            =   30
            TabIndex        =   11
            Top             =   45
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   9499
            _LayoutType     =   0
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   -1  'True
            Splits(0)._GSX_SAVERECORDSELECTORS=   0
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   0
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            MultiSelect     =   2
            DeadAreaBackColor=   14215660
            RowDividerColor =   14215660
            RowSubDividerColor=   14215660
            DirectionAfterEnter=   1
            DirectionAfterTab=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=126,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
            _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
            _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
            _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(38)  =   "Named:id=33:Normal"
            _StyleDefs(39)  =   ":id=33,.parent=0"
            _StyleDefs(40)  =   "Named:id=34:Heading"
            _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(42)  =   ":id=34,.wraptext=-1"
            _StyleDefs(43)  =   "Named:id=35:Footing"
            _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(45)  =   "Named:id=36:Selected"
            _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(47)  =   "Named:id=37:Caption"
            _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(49)  =   "Named:id=38:HighlightRow"
            _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(51)  =   "Named:id=39:EvenRow"
            _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(53)  =   "Named:id=40:OddRow"
            _StyleDefs(54)  =   ":id=40,.parent=33"
            _StyleDefs(55)  =   "Named:id=41:RecordSelector"
            _StyleDefs(56)  =   ":id=41,.parent=34"
            _StyleDefs(57)  =   "Named:id=42:FilterBar"
            _StyleDefs(58)  =   ":id=42,.parent=33"
         End
      End
   End
   Begin MSMask.MaskEdBox MaskCari 
      Height          =   270
      Left            =   4950
      TabIndex        =   10
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
Attribute VB_Name = "frmCaller1"
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
Tampil
txtCari.Text = ""
txtCari.SetFocus
Err.Clear
End Sub

Private Sub Command1_Click()
Call dgDetail_DblClick
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

Dim row As Variant
If dgDetail.SelBookmarks.Count > 1 Then
   For Each row In dgDetail.SelBookmarks
      rsSiswa.Bookmark = row
      CollectingData rsSiswa.Fields("nis")
      
   Next
Else
   dgDetail_DblClick
End If

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
    Tampil
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






