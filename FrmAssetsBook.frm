VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{43E6F32B-2B03-46D3-9276-69426FE6D51B}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmAssetsBook 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assets Book"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   Tag             =   "Assets Book"
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5985
      Left            =   75
      ScaleHeight     =   5955
      ScaleWidth      =   10335
      TabIndex        =   36
      Tag             =   "Class Setup"
      Top             =   60
      Width           =   10365
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   5205
         Left            =   105
         ScaleHeight     =   5175
         ScaleWidth      =   10095
         TabIndex        =   37
         Top             =   285
         Width           =   10125
         Begin MSDataListLib.DataCombo DataCombo1 
            DataField       =   "AssetID"
            Height          =   315
            Left            =   2160
            TabIndex        =   1
            Tag             =   "ASM"
            Top             =   195
            Width           =   3420
            _ExtentX        =   6033
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Description"
            BoundColumn     =   "AssetID"
            Text            =   ""
         End
         Begin VB.ComboBox Combo1 
            DataField       =   "AmzCode"
            Height          =   315
            Index           =   3
            ItemData        =   "FrmAssetsBook.frx":0000
            Left            =   6615
            List            =   "FrmAssetsBook.frx":0019
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Tag             =   "ASM"
            Top             =   2445
            Width           =   3375
         End
         Begin VB.ComboBox Combo1 
            DataField       =   "SwitchOver"
            Height          =   315
            Index           =   2
            ItemData        =   "FrmAssetsBook.frx":005A
            Left            =   2160
            List            =   "FrmAssetsBook.frx":0064
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Tag             =   "ASM"
            Top             =   4695
            Width           =   3375
         End
         Begin VB.ComboBox Combo1 
            DataField       =   "AvgConv"
            Height          =   315
            Index           =   1
            ItemData        =   "FrmAssetsBook.frx":0082
            Left            =   2160
            List            =   "FrmAssetsBook.frx":00AD
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Tag             =   "ASM"
            Top             =   4365
            Width           =   3375
         End
         Begin VB.ComboBox Combo1 
            DataField       =   "DeprMethod"
            Height          =   315
            Index           =   0
            ItemData        =   "FrmAssetsBook.frx":017F
            Left            =   2160
            List            =   "FrmAssetsBook.frx":01AD
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Tag             =   "ASM"
            Top             =   4050
            Width           =   3375
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "PlaceServiceDate"
            Height          =   315
            Index           =   0
            Left            =   2160
            TabIndex        =   5
            Tag             =   "ASM"
            Top             =   855
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MMM/yyyy"
            Format          =   58261507
            CurrentDate     =   38611
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "BegYearCost"
            Height          =   285
            Index           =   0
            Left            =   2160
            TabIndex        =   9
            Tag             =   "ASM"
            Top             =   1530
            Width           =   3375
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "CostBasis"
            Height          =   285
            Index           =   1
            Left            =   2160
            TabIndex        =   13
            Tag             =   "ASM"
            Top             =   1845
            Width           =   3375
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "SalvageVal"
            Height          =   285
            Index           =   2
            Left            =   2160
            TabIndex        =   17
            Tag             =   "ASM"
            Top             =   2145
            Width           =   3375
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "YearlyDeprRate"
            Height          =   285
            Index           =   3
            Left            =   2160
            TabIndex        =   19
            Tag             =   "ASM"
            Top             =   2460
            Width           =   3375
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "CurrentDepr"
            Height          =   285
            Index           =   4
            Left            =   2160
            TabIndex        =   23
            Tag             =   "ASM"
            Top             =   2775
            Width           =   3375
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "DeprYTD"
            Height          =   285
            Index           =   5
            Left            =   2160
            TabIndex        =   27
            Tag             =   "ASM"
            Top             =   3090
            Width           =   3375
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "DeprLTD"
            Height          =   285
            Index           =   6
            Left            =   2160
            TabIndex        =   29
            Tag             =   "ASM"
            Top             =   3405
            Width           =   3375
         End
         Begin VB.TextBox txtFields 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "OrgLifeYear"
            Height          =   285
            Index           =   7
            Left            =   6645
            TabIndex        =   11
            Tag             =   "ASM"
            Top             =   1530
            Width           =   3375
         End
         Begin VB.TextBox txtFields 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "OrgLifeDay"
            Height          =   285
            Index           =   8
            Left            =   6645
            TabIndex        =   15
            Tag             =   "ASM"
            Top             =   1845
            Width           =   3375
         End
         Begin VB.TextBox txtFields 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "AmzAmount"
            Height          =   285
            Index           =   10
            Left            =   6645
            TabIndex        =   25
            Tag             =   "ASM"
            Top             =   2775
            Width           =   3375
         End
         Begin VB.TextBox txtFields 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "BookID"
            Height          =   285
            Index           =   15
            Left            =   2160
            TabIndex        =   3
            Tag             =   "ASM"
            Top             =   540
            Width           =   2190
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "DeprecDate"
            Height          =   315
            Index           =   1
            Left            =   2160
            TabIndex        =   7
            Tag             =   "ASM"
            Top             =   1185
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MMM/yyyy"
            Format          =   58261507
            CurrentDate     =   38611
         End
         Begin VB.Line Line1 
            Index           =   17
            X1              =   240
            X2              =   2310
            Y1              =   4995
            Y2              =   4995
         End
         Begin VB.Line Line1 
            Index           =   16
            X1              =   240
            X2              =   2310
            Y1              =   4665
            Y2              =   4665
         End
         Begin VB.Line Line1 
            Index           =   15
            X1              =   240
            X2              =   2310
            Y1              =   4335
            Y2              =   4335
         End
         Begin VB.Line Line1 
            Index           =   14
            X1              =   5595
            X2              =   7665
            Y1              =   3045
            Y2              =   3045
         End
         Begin VB.Line Line1 
            Index           =   13
            X1              =   5595
            X2              =   7665
            Y1              =   2745
            Y2              =   2745
         End
         Begin VB.Line Line1 
            Index           =   12
            X1              =   5595
            X2              =   7665
            Y1              =   2115
            Y2              =   2115
         End
         Begin VB.Line Line1 
            Index           =   11
            X1              =   5595
            X2              =   7665
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line Line1 
            Index           =   10
            X1              =   240
            X2              =   2310
            Y1              =   3675
            Y2              =   3675
         End
         Begin VB.Line Line1 
            Index           =   9
            X1              =   240
            X2              =   2310
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   240
            X2              =   2310
            Y1              =   3045
            Y2              =   3045
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   240
            X2              =   2310
            Y1              =   2730
            Y2              =   2730
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   240
            X2              =   2310
            Y1              =   2415
            Y2              =   2415
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   240
            X2              =   2310
            Y1              =   2115
            Y2              =   2115
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   240
            X2              =   2310
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   240
            X2              =   2310
            Y1              =   1485
            Y2              =   1485
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   240
            X2              =   2310
            Y1              =   1155
            Y2              =   1155
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   240
            X2              =   2310
            Y1              =   810
            Y2              =   810
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   240
            X2              =   2325
            Y1              =   495
            Y2              =   495
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Begin Year Cost"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   8
            Top             =   1575
            Width           =   1140
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cost Basis"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   12
            Top             =   1890
            Width           =   735
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Salvage Value"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   16
            Top             =   2190
            Width           =   1005
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Yearly Depr. Rate"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   18
            Top             =   2505
            Width           =   1290
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Depriciation"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   22
            Top             =   2820
            Width           =   1440
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "YTD Depreciation"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   26
            Top             =   3135
            Width           =   1230
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LTD Depreciation"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   28
            Top             =   3450
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OrgLifeYear"
            Height          =   195
            Index           =   7
            Left            =   5595
            TabIndex        =   10
            Top             =   1575
            Width           =   855
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OrgLifeDay"
            Height          =   195
            Index           =   8
            Left            =   5595
            TabIndex        =   14
            Top             =   1890
            Width           =   810
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AmzCode"
            Height          =   195
            Index           =   9
            Left            =   5595
            TabIndex        =   20
            Top             =   2505
            Width           =   675
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AmzAmount"
            Height          =   195
            Index           =   10
            Left            =   5595
            TabIndex        =   24
            Top             =   2820
            Width           =   855
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Depreciation Method"
            Height          =   195
            Index           =   11
            Left            =   240
            TabIndex        =   30
            Top             =   4110
            Width           =   1485
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Averaging Convention"
            Height          =   195
            Index           =   12
            Left            =   240
            TabIndex        =   32
            Top             =   4425
            Width           =   1605
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SwitchOver"
            Height          =   195
            Index           =   13
            Left            =   240
            TabIndex        =   34
            Top             =   4740
            Width           =   825
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AssetID"
            Height          =   195
            Index           =   14
            Left            =   240
            TabIndex        =   0
            Top             =   270
            Width           =   570
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BookID"
            Height          =   195
            Index           =   15
            Left            =   240
            TabIndex        =   2
            Top             =   585
            Width           =   510
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PlaceIn Service Date"
            Height          =   195
            Index           =   16
            Left            =   240
            TabIndex        =   4
            Top             =   915
            Width           =   1485
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Depreciated To Date"
            Height          =   195
            Index           =   17
            Left            =   240
            TabIndex        =   6
            Top             =   1245
            Width           =   1485
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   38
      Top             =   6090
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
      LimitRecordData =   "36"
   End
End
Attribute VB_Name = "FrmAssetsBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RcGroup As New DBQuick

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub DataCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub DTPicker1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Form_Load()
HiasForm Picture1, Me
CenterForm Picture2, Me
RcGroup.DBOpen "SELECT [No Aktiva] AS [AssetID], [Nama Aktiva] AS [Description] FROM         [Tabel Aktiva Tetap] ORDER BY [No Aktiva]", CNN, lckLockBatch
Set DataCombo1.RowSource = RcGroup.DBRecordset
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmAssetsBook
    .BindFormTAG = "ASM"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT     AssetID, BookID, PlaceServiceDate, DeprecDate, BegYearCost, CostBasis, SalvageVal, YearlyDeprRate, CurrentDepr, DeprYTD, DeprLTD, OrgLifeYear, OrgLifeDay, AmzCode, AmzAmount, DeprMethod, AvgConv, SwitchOver FROM  AssetBook ORDER BY BookID"
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
RcGroup.CloseDB
Set RcGroup = Nothing
End Sub

Private Sub Form_Resize()
'
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmAssetsBook = Nothing
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
MyDDE.PrepareAppend = " INSERT INTO AssetBook" & _
                      " (AssetID, BookID, PlaceServiceDate, DeprecDate, BegYearCost, CostBasis, SalvageVal, YearlyDeprRate, CurrentDepr, DeprYTD, DeprLTD, OrgLifeYear, OrgLifeDay, AmzCode, AmzAmount, DeprMethod, AvgConv, SwitchOver)" & _
                      " VALUES (N'" & DataCombo1.BoundText & "', N'" & txtFields(15) & "' , convert(datetime,'" & Format(DTPicker1(0).value, "dd/mm/yy") & "',3), convert(datetime,'" & Format(DTPicker1(1).value, "dd/mm/yy") & "',3), " & CDbl(txtFields(0)) & ", " & CDbl(txtFields(1)) & ", " & CDbl(txtFields(2)) & ", " & _
                      CDbl(txtFields(3)) & ", " & CDbl(txtFields(4)) & ", " & CDbl(txtFields(5)) & "," & CDbl(txtFields(6)) & "," & CDbl(txtFields(7)) & "," & CDbl(txtFields(8)) & ",N'" & Combo1(3) & "'," & CDbl(txtFields(10)) & ",N'" & Combo1(0) & "',N'" & Combo1(1) & "',N'" & Combo1(2) & "')"
                      
MyDDE.PrepareUpdate = " UPDATE AssetBook " & _
                      " SET  [PlaceServiceDate] =convert(datetime,'" & Format(DTPicker1(0).value, "dd/mm/yy") & "',3), DeprecDate = Convert(Datetime,'" & Format(DTPicker1(0).value, "dd/mm/yy") & "',3)" & _
                      " , BegYearCost = " & CDbl(txtFields(0)) & "," & _
                      " CostBasis = " & CDbl(txtFields(1)) & "," & _
                      " SalvageVal = " & CDbl(txtFields(2)) & "," & _
                      " YearlyDeprRate = " & CDbl(txtFields(3)) & "," & _
                      " CurrentDepr = " & CDbl(txtFields(4)) & "," & _
                      " DeprYTD = " & CDbl(txtFields(5)) & "," & _
                      " DeprLTD = " & CDbl(txtFields(6)) & "," & _
                      " OrgLifeYear = " & CDbl(txtFields(7)) & "," & _
                      " OrgLifeDay = " & CDbl(txtFields(8)) & "," & _
                      " AmzCode = N'" & Combo1(3) & "'," & _
                      " AmzAmount = " & CDbl(txtFields(10)) & "," & _
                      " DeprMethod = N'" & Combo1(0) & "'," & _
                      " AvgConv = N'" & Combo1(1) & "'," & _
                      " SwitchOver = N'" & Combo1(2) & "'" & _
                      " WHERE ([BookID]= N'" & txtFields(15) & "')"
MyDDE.PrepareDelete = " DELETE FROM AssetBook WHERE  ([BookID] = N'" & txtFields(15) & "')"
Err.Clear
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbAddNew:
               MyDDE.IsChildMemberReady = True
               NormalNol
               DataCombo1.SetFocus
       Case tmbEdit:
               DataCombo1.Enabled = False
               txtFields(15).Enabled = False
               DTPicker1(0).SetFocus
       Case tmbDelete:
            If MyDDE.IsChildMemberReady = True Then
            End If
End Select
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
            Else
               MyDDE.IsChildMemberReady = False
            End If

End Select
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Picture1.Parent.hwnd
End Sub

Private Sub txtFields_Change(Index As Integer)
If txtFields(Index) = "" Then txtFields(Index) = 0
End Sub

Private Sub txtFields_GotFocus(Index As Integer)
Block txtFields(Index)
End Sub

Private Sub txtFields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
ValidNum KeyAscii
End Sub

Private Sub NormalNol()
txtFields(0) = 0
txtFields(1) = 0
txtFields(2) = 0
txtFields(3) = 0
txtFields(4) = 0
txtFields(5) = 0
txtFields(6) = 0
txtFields(7) = 0
txtFields(8) = 0
txtFields(10) = 0
txtFields(15) = 0
End Sub
