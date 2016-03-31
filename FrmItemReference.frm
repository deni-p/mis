VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmItemReference 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Substitusi Stok"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9885
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmItemReference.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   Tag             =   "Inventory Reference"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   5370
      Left            =   0
      ScaleHeight     =   5370
      ScaleWidth      =   9885
      TabIndex        =   8
      Top             =   0
      Width           =   9885
      Begin TabDlg.SSTab SSTab1 
         Height          =   5070
         Left            =   60
         TabIndex        =   9
         Top             =   120
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   8943
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   15380335
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "List"
         TabPicture(0)   =   "FrmItemReference.frx":6852
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Picture3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Detail"
         TabPicture(1)   =   "FrmItemReference.frx":686E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Picture4"
         Tab(1).ControlCount=   1
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00EAAF6F&
            Height          =   4545
            Left            =   -74895
            ScaleHeight     =   4485
            ScaleWidth      =   9525
            TabIndex        =   12
            Top             =   405
            Width           =   9585
            Begin VB.TextBox txtBox 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "Last Cost"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0;(#,##0)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   330
               Index           =   5
               Left            =   1410
               MaxLength       =   13
               ScrollBars      =   2  'Vertical
               TabIndex        =   4
               Tag             =   "Partner"
               Text            =   "Text1"
               Top             =   1110
               Width           =   2250
            End
            Begin VB.TextBox txtBox 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "Average Cost"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0;(#,##0)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   330
               Index           =   4
               Left            =   6645
               Locked          =   -1  'True
               MaxLength       =   13
               ScrollBars      =   2  'Vertical
               TabIndex        =   6
               Tag             =   "Partner"
               Text            =   "Text1"
               Top             =   765
               Width           =   2250
            End
            Begin VB.TextBox txtBox 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "Fixed Cost"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0;(#,##0)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   330
               Index           =   3
               Left            =   1410
               Locked          =   -1  'True
               MaxLength       =   13
               ScrollBars      =   2  'Vertical
               TabIndex        =   3
               Tag             =   "Partner"
               Text            =   "Text1"
               Top             =   765
               Width           =   2250
            End
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "Nama Barang"
               Height          =   330
               Index           =   2
               Left            =   1410
               MaxLength       =   50
               TabIndex        =   2
               Tag             =   "Partner"
               Text            =   "Text1"
               Top             =   420
               Width           =   3945
            End
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "UOM"
               Height          =   330
               Index           =   1
               Left            =   6645
               MaxLength       =   15
               ScrollBars      =   2  'Vertical
               TabIndex        =   5
               Tag             =   "Partner"
               Text            =   "Text1"
               Top             =   420
               Width           =   2250
            End
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "Kode Barang"
               Height          =   330
               Index           =   0
               Left            =   1410
               MaxLength       =   15
               TabIndex        =   1
               Tag             =   "Partner"
               Text            =   "Text1"
               Top             =   75
               Width           =   2250
            End
            Begin TabDlg.SSTab SSTab2 
               Height          =   2865
               Left            =   165
               TabIndex        =   7
               Top             =   1545
               Width           =   8925
               _ExtentX        =   15743
               _ExtentY        =   5054
               _Version        =   393216
               Style           =   1
               Tabs            =   2
               Tab             =   1
               TabsPerRow      =   2
               TabHeight       =   520
               BackColor       =   15380335
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TabCaption(0)   =   "Alternate"
               TabPicture(0)   =   "FrmItemReference.frx":688A
               Tab(0).ControlEnabled=   0   'False
               Tab(0).Control(0)=   "Picture5"
               Tab(0).ControlCount=   1
               TabCaption(1)   =   "Supplier"
               TabPicture(1)   =   "FrmItemReference.frx":68A6
               Tab(1).ControlEnabled=   -1  'True
               Tab(1).Control(0)=   "Picture7"
               Tab(1).Control(0).Enabled=   0   'False
               Tab(1).ControlCount=   1
               Begin VB.PictureBox Picture7 
                  Height          =   2400
                  Left            =   75
                  ScaleHeight     =   2340
                  ScaleWidth      =   8700
                  TabIndex        =   15
                  Top             =   375
                  Width           =   8760
                  Begin MSDataGridLib.DataGrid DataGrid1 
                     Bindings        =   "FrmItemReference.frx":68C2
                     Height          =   2325
                     Index           =   1
                     Left            =   0
                     TabIndex        =   16
                     Top             =   0
                     Width           =   8670
                     _ExtentX        =   15293
                     _ExtentY        =   4101
                     _Version        =   393216
                     AllowUpdate     =   -1  'True
                     Appearance      =   0
                     BorderStyle     =   0
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
                     ColumnCount     =   5
                     BeginProperty Column00 
                        DataField       =   "Nama Perusahaan"
                        Caption         =   "Nama Perusahaan"
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
                     BeginProperty Column01 
                        DataField       =   "Ref Code"
                        Caption         =   "Ref Code"
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
                     BeginProperty Column02 
                        DataField       =   "Keterangan"
                        Caption         =   "Keterangan"
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
                        Caption         =   "UOM"
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
                        DataField       =   "QTY"
                        Caption         =   "QTY"
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
                        EndProperty
                        BeginProperty Column01 
                        EndProperty
                        BeginProperty Column02 
                        EndProperty
                        BeginProperty Column03 
                        EndProperty
                        BeginProperty Column04 
                        EndProperty
                     EndProperty
                  End
               End
               Begin VB.PictureBox Picture5 
                  Height          =   2400
                  Left            =   -74925
                  ScaleHeight     =   2340
                  ScaleWidth      =   8700
                  TabIndex        =   13
                  Top             =   375
                  Width           =   8760
                  Begin MSDataGridLib.DataGrid DataGrid1 
                     Bindings        =   "FrmItemReference.frx":68D7
                     Height          =   2325
                     Index           =   0
                     Left            =   0
                     TabIndex        =   14
                     Top             =   0
                     Width           =   8670
                     _ExtentX        =   15293
                     _ExtentY        =   4101
                     _Version        =   393216
                     AllowUpdate     =   -1  'True
                     Appearance      =   0
                     BorderStyle     =   0
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
                     ColumnCount     =   4
                     BeginProperty Column00 
                        DataField       =   "Alternate ID"
                        Caption         =   "Alternate ID"
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
                        DataField       =   "Keterangan"
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
                     BeginProperty Column02 
                        DataField       =   "UOM"
                        Caption         =   "UOM"
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
                        DataField       =   "QTY"
                        Caption         =   "QTY"
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
                        BeginProperty Column02 
                        EndProperty
                        BeginProperty Column03 
                        EndProperty
                     EndProperty
                  End
               End
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nama Barang"
               Height          =   195
               Index           =   12
               Left            =   225
               TabIndex        =   29
               Top             =   488
               Width           =   960
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kode Barang"
               Height          =   195
               Index           =   11
               Left            =   225
               TabIndex        =   28
               Top             =   143
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fixed Cost"
               Height          =   195
               Index           =   10
               Left            =   225
               TabIndex        =   27
               Top             =   833
               Width           =   765
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Last Cost"
               Height          =   195
               Index           =   9
               Left            =   210
               TabIndex        =   26
               Top             =   1178
               Width           =   675
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "UOM"
               Height          =   195
               Index           =   8
               Left            =   5475
               TabIndex        =   25
               Top             =   488
               Width           =   345
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Average Cost"
               Height          =   195
               Index           =   7
               Left            =   5475
               TabIndex        =   24
               Top             =   833
               Width           =   990
            End
            Begin VB.Line Line1 
               Index           =   2
               X1              =   195
               X2              =   1620
               Y1              =   735
               Y2              =   735
            End
            Begin VB.Line Line1 
               Index           =   0
               X1              =   195
               X2              =   1620
               Y1              =   390
               Y2              =   390
            End
            Begin VB.Line Line1 
               Index           =   4
               X1              =   5475
               X2              =   6900
               Y1              =   735
               Y2              =   735
            End
            Begin VB.Line Line1 
               Index           =   1
               X1              =   5475
               X2              =   6900
               Y1              =   1080
               Y2              =   1080
            End
            Begin VB.Line Line1 
               Index           =   3
               X1              =   195
               X2              =   1620
               Y1              =   1425
               Y2              =   1425
            End
            Begin VB.Line Line1 
               Index           =   5
               X1              =   195
               X2              =   1620
               Y1              =   1080
               Y2              =   1080
            End
         End
         Begin VB.PictureBox Picture3 
            Height          =   4575
            Left            =   105
            ScaleHeight     =   4515
            ScaleWidth      =   9510
            TabIndex        =   10
            Top             =   390
            Width           =   9570
            Begin MSComctlLib.ListView ListView1 
               Height          =   4515
               Left            =   0
               TabIndex        =   11
               Top             =   0
               Width           =   9510
               _ExtentX        =   16775
               _ExtentY        =   7964
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               AllowReorder    =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   15380335
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   6
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Kode Barang"
                  Object.Width           =   2910
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Nama Barang"
                  Object.Width           =   5644
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "UOM"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   3
                  Text            =   "Fixed Cost"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   4
                  Text            =   "Average Cost"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   5
                  Text            =   "Last Cost"
                  Object.Width           =   2540
               EndProperty
            End
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5370
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Barang"
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   23
      Top             =   330
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Barang"
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fixed Cost"
      Height          =   195
      Index           =   4
      Left            =   0
      TabIndex        =   21
      Top             =   660
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Cost"
      Height          =   195
      Index           =   2
      Left            =   0
      TabIndex        =   20
      Top             =   990
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UOM"
      Height          =   195
      Index           =   5
      Left            =   5205
      TabIndex        =   19
      Top             =   345
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Average Cost"
      Height          =   195
      Index           =   6
      Left            =   5205
      TabIndex        =   18
      Top             =   675
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cost Rollup"
      Height          =   195
      Index           =   3
      Left            =   5205
      TabIndex        =   17
      Top             =   1005
      Width           =   810
   End
End
Attribute VB_Name = "FrmItemReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsMytr                         As New DBQuick
Private RcAlter                         As New DBQuick
Private RcCompli                        As New DBQuick
Private RcCust                          As New DBQuick
Private RcPart                          As New DBQuick
Private RcPartner                       As New DBQuick
Private WithEvents mCall                As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private MyData                          As New clsTransaksi
Private MEdit As Boolean
Private mFirstCaller             As Boolean
Private mAccount                        As String
Private mKeyLoad                        As Boolean

Private Sub DataGrid1_ButtonClick(Index As Integer, ByVal ColIndex As Integer)
If Index = 1 Then
   If DataGrid1(Index).col = 0 Then
      OpenDetailPartner 1
   ElseIf DataGrid1(Index).col = 1 Then
      OpenDetailPartner 2
   End If
End If
End Sub

Private Sub DataGrid1_Error(Index As Integer, ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub

Private Sub DataGrid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Form_KeyDown KeyCode, Shift
'ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub DataGrid1_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
If MEdit = True Then
   If Index = 0 Then
        If DataGrid1(Index).col = 3 Then
           DataGrid1(Index).MarqueeStyle = dbgFloatingEditor
           DataGrid1(Index).AllowUpdate = True
        Else
           DataGrid1(Index).AllowUpdate = False
           DataGrid1(Index).MarqueeStyle = dbgHighlightRow
        End If
   Else
        If DataGrid1(Index).col <= 1 Or DataGrid1(Index).col = 4 Then
           DataGrid1(Index).MarqueeStyle = dbgFloatingEditor
           DataGrid1(Index).AllowUpdate = True
        Else
           DataGrid1(Index).AllowUpdate = False
           DataGrid1(Index).MarqueeStyle = dbgHighlightRow
        End If
   End If
Else
   DataGrid1(Index).MarqueeStyle = dbgHighlightRow
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If mKeyLoad = False Then mKeyLoad = True Else mKeyLoad = False
If mKeyLoad = False Then ScanKey KeyCode, Shift, MyDDE
'Call DGPurchase_KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_Load()
GridLayout
HiasFormManTell Picture2, Me
MyDDE.SetPermissions = aksess.MayDo("Substitusi")
'HiasForm Picture1, Me
With MyDDE
     .EditModeReplace = False
     .SetPermissions = UserAddnewDenied
     Set .BindForm = FrmItemReference
     .BindFormTAG = "Partner"
     Set .ActiveConnection = CNN
     .PrepareQuery = " SELECT NoItem AS [Kode Barang], ItemName AS [Nama Barang], UOM, FixCost AS [Fixed Cost], AvgCost AS [Average Cost], LastCost AS [Last Cost] FROM         Inventory WHERE     (Manufacture = 0)"
End With
SSTab1.Tab = 0
SSTab2.Tab = 0
OpenHeader
Set mCall = New frmCaller
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set MyData = Nothing
MyDDE.ClearRecordset
clsMytr.CloseDB
Set mCall = Nothing
End Sub

Private Sub Form_Resize()

Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmItemReference = Nothing
End Sub

'Private Sub ListView1_Click()
'messagebox ListView1.SelectedItem.Index
'End Sub

Private Sub ListView1_LostFocus()
If SSTab1.Tab = 0 Then MyDDE.SetFocus
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub


Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit, tmbDelete:
             MyDDE.CancelTrans = False
       Case tmbDetail:
            MyDDE.CancelTrans = mFirstCaller
            If MyData.CheckGridKosong(MyDDE.ChildRecordset) = True Then
               MyDDE.CancelTrans = True
               MessageBox "Data transaksi belum lengkap." & "Silahkan dicek kembali.", "Peringatan", msgOkOnly, msgCrtical
            Else
               MyDDE.CancelTrans = mFirstCaller
            End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If MyData.CheckGridKosong(RcAlter.DBRecordset) = True Or MyData.CheckGridKosong(RcPart.DBRecordset) = True Then
                  MyDDE.CancelTrans = True
                  MessageBox "Data transaksi belum lengkap." & "Silahkan dicek kembali.", "Peringatan", msgOkOnly, msgCrtical
               Else
                  MyDDE.CancelTrans = mFirstCaller
                  MyDDE.IsChildMemberReady = True
                  PrepareQuery
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo xErr
Select Case AdReasonActiveDb
       Case tmbEdit:
            MEdit = True
            txtBox(0).Enabled = False
            txtBox(1).Enabled = False
            txtBox(2).Enabled = False
            txtBox(3).Enabled = False
            txtBox(4).Enabled = False
            txtBox(5).Enabled = False
            If SSTab2.Tab = 0 Then
               DataGrid1(0).SetFocus
            Else
               DataGrid1(1).SetFocus
            End If
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
                With RcAlter.DBRecordset
                     If .Recordcount <> 0 Then
                     .MoveFirst
                     If SendDataToServer(" Delete From  [Inventory Alternates] WHERE     (NoItem = N'" & txtBox(0) & "') AND (TypeItem=0) ") = True Then
                     Do
                       If RcAlter.DBRecordset.EOF Then Exit Do
                          SendDataToServer " INSERT INTO [Inventory Alternates]" & _
                                           " ( NoItem, AlternateID, Description,UOM,QTY,TypeItem)" & _
                                           " VALUES (N'" & txtBox(0) & "', N'" & .Fields("Alternate ID") & "', N'" & .Fields("Keterangan") & "', N'" & .Fields("UOM") & "'," & CDbl(.Fields("QTY")) & ",0)"
                     .MoveNext
                     Loop
                     End If
                     .MoveLast
                     End If
                End With
                
                With RcPart.DBRecordset
                     If .Recordcount <> 0 Then
                     .MoveFirst
                     If SendDataToServer(" Delete From  [Inventory Partner Alternates] WHERE     (AlternateID = N'" & txtBox(0) & "')  ") = True Then
                     Do
                       If .EOF Then Exit Do
                          SendDataToServer " INSERT INTO [Inventory Partner Alternates]" & _
                                           " (AlternateID, PartnerID, RefCode, [DESC], UOM, QTY)" & _
                                           " VALUES (N'" & txtBox(0) & "', N'" & .Fields("ID") & "', N'" & .Fields("Ref Code") & "', N'" & .Fields("Keterangan") & "', N'" & .Fields("UOM") & "', " & CDbl(.Fields("QTY")) & ")"
                     .MoveNext
                     Loop
                     End If
                     .MoveLast
                     End If
                End With
                MEdit = False
            End If
       Case tmbCancel:
            If MyDDE.ChildRecordset.Recordcount = 0 Then
               'OpenAlternate txtBox(0)
               MEdit = False
'               DGPurchase.Columns(6).Visible = True
'               DGPurchase.Columns(7).Visible = False
'               If Me.Caption = "P.O Transaksi" Then chkPo.Enabled = False
             Else
'               DGPurchase.Columns(6).Visible = False
'               DGPurchase.Columns(7).Visible = True
               MEdit = True
             End If
       Case tmbDetail:
            If mFirstCaller = False Then
               OpenDetailPartner SSTab2.Tab
               MEdit = True
            End If
       Case tmbPrint:
            CallRPTReport "Item Reference Table.rpt", "Select * From [Item Reference Table] where [Kode Barang] ='" & txtBox(0) & "'"
       Case tmbQuit:
            Unload Me
            Set MyDDE.BindForm = Nothing
End Select
DataGrid1(1).Columns(0).Button = MEdit
DataGrid1(1).Columns(1).Button = MEdit
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenAlternate MyDDE.GetFieldByName("Kode Barang")
OpenPart MyDDE.GetFieldByName("Kode Barang")
End Sub

Private Sub OpenDetailPartner(ByVal Index As Integer)
On Error GoTo Hell:
Select Case Index
       Case 0:
            RcPartner.DBOpen "SELECT     NoItem AS [Item Alternate], ItemName AS Keterangan, UOM FROM         Inventory WHERE     (Manufacture = 0) AND (NoItem <> N'" & txtBox(0) & "') ORDER BY NoItem", CNN, lckLockReadOnly
            mFirstCaller = True
       Case 1: RcPartner.DBOpen "SELECT     PartnerID AS ID, CompanyName AS [Nama Perusahaan] FROM         PartnerDB WHERE     (PartnerType = N'SUPPLIER') ORDER BY PartnerID", CNN, lckLockReadOnly
       Case 2: RcPartner.DBOpen "SELECT     NoItem AS [Kode Barang], ItemName AS [Nama Barang], UOM FROM         Inventory", CNN, lckLockReadOnly
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 0: mCall.FromTagActive = "INVENTORY ALTERNATE"
          Case 1:
               mCall.FromTagActive = "MASTER SUPPLIER"
               DataGrid1(1).Columns(0).Button = True
               DataGrid1(1).Columns(1).Button = True
          Case 2: mCall.FromTagActive = "INVENTORY"
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me
Else
   MessageBox "Data Belum Ada.", "Peringatan", msgOkOnly, msgExclamation
   DataGrid1(1).Columns(0).Button = False
   DataGrid1(1).Columns(1).Button = False
End If
Exit Sub
Hell:
'    messagebox Err.Description
    Err.Clear
End Sub

Private Sub PrepareQuery()
On Error GoTo xErr
With MyDDE
    .PrepareAppend = " INSERT INTO [Inventory]" & _
                     " (FixCost)" & _
                     " VALUES (" & CDbl(txtBox(3)) & ")"
'MessageBox .PrepareAppend
    .PrepareUpdate = " UPDATE [Inventory] Set FixCost=" & CDbl(txtBox(3)) & ",AvgCost=" & CDbl(txtBox(4)) & ",LastCost=" & CDbl(txtBox(5)) & " WHERE     ([NoItem] = N'" & txtBox(0) & "')"

    .PrepareDelete = " DELETE FROM [Inventory] WHERE   ([NoItem] = N'" & txtBox(0) & "') "
End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub ListView1_DblClick()
If MyDDE.ActiveRecordset.Recordcount <> 0 Then
   SSTab1.Tab = 1
   SSTab2.Tab = 0
End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
If MyDDE.ActiveRecordset.Recordcount <> 0 Then
   MyDDE.FindStringData "[Kode Barang]='" & Item.Text & "'"
End If
End Sub

Private Sub mCall_BeforeUnload()
Select Case mCall.FromTagActive
       Case "INVENTORY ALTERNATE":
            If FindOwnRecordset(MyDDE.ChildRecordset, "[Alternate ID] = '" & MyDDE.ChildRecordset.Fields("Alternate ID") & "'") = True Then
               MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Alternate ID") & " Sudah Ada....! Silahkan Diulangi", "Peringatan", msgOkOnly, msgCrtical
               MyDDE.ChildRecordset.CancelBatch adAffectCurrent
               If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
            Else
               If Not IsNull(MyDDE.ChildRecordset.Fields(0)) = True Or MyDDE.ChildRecordset.Fields(0) = "" Then
                  'If MyDDE.ChildRecordset.Fields(0) = "" Then
'                     MyDDE.ChildRecordset.CancelBatch adAffectCurrent
'                     If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
                Else
                     MyDDE.ChildRecordset.CancelBatch adAffectCurrent
                     If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
                  'End If
               End If
            End If
            mFirstCaller = False
       Case "MASTER SUPPLIER":
            If FindOwnRecordset(MyDDE.ChildRecordset, "[ID] = '" & MyDDE.ChildRecordset.Fields("ID") & "'") = True Then
               MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Nama Perusahaan") & " Sudah Ada....! Silahkan Diulangi", "Peringatan", msgOkOnly, msgCrtical
               MyDDE.ChildRecordset.CancelBatch adAffectCurrent
               If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
            Else
               If Not IsNull(MyDDE.ChildRecordset.Fields(0)) = True Or MyDDE.ChildRecordset.Fields(0) = "" Then
                  'If MyDDE.ChildRecordset.Fields(0) = "" Then
'                     MyDDE.ChildRecordset.CancelBatch adAffectCurrent
'                     If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
                Else
                     MyDDE.ChildRecordset.CancelBatch adAffectCurrent
                     If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
                  'End If
               End If
            End If
            mFirstCaller = False
End Select
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case mCall.FromTagActive
       Case "INVENTORY ALTERNATE":
                     MyDDE.ChildRecordset.Fields("Alternate ID") = mCall.GetFieldByName(0)
                     MyDDE.ChildRecordset.Fields("Keterangan") = mCall.GetFieldByName(1)
                     MyDDE.ChildRecordset.Fields("uom") = mCall.GetFieldByName(2)
                     MyDDE.ChildRecordset.Fields("QTY") = 1
       Case "MASTER SUPPLIER":
                With RcCompli.DBRecordset
                     MyDDE.ChildRecordset.Fields("ID") = mCall.GetFieldByName(0)
                     MyDDE.ChildRecordset.Fields("Nama Perusahaan") = mCall.GetFieldByName(1)
                     CariDataDefaultItem MyDDE.ChildRecordset.Fields("ID")
                End With
       Case "INVENTORY":
                With RcCompli.DBRecordset
                     MyDDE.ChildRecordset.Fields("Ref Code") = mCall.GetFieldByName(0)
                     MyDDE.ChildRecordset.Fields("Keterangan") = mCall.GetFieldByName(1)
                     MyDDE.ChildRecordset.Fields("uom") = mCall.GetFieldByName(2)
                     MyDDE.ChildRecordset.Fields("QTY") = 1
                End With
End Select
End Sub

Private Sub OpenHeader()
Dim Rc As New DBQuick
Dim I As Integer
Dim Avdata As Variant
Set Rc.DBRecordset = MyDDE.ActiveRecordset.Clone(adLockReadOnly)
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            With ListView1.ListItems.Add(, , Avdata(0, I))
                 .SubItems(1) = Avdata(1, I)
                 .SubItems(2) = Avdata(2, I)
                 .SubItems(3) = FormatNumber(Avdata(3, I), 0)
                 .SubItems(4) = FormatNumber(Avdata(4, I), 0)
                 .SubItems(5) = FormatNumber(Avdata(5, I), 0)
            End With
        Next I
     Else
     End If
End With
End Sub

Private Sub OpenAlternate(ByVal Param As String)
RcAlter.DBOpen "SELECT     [Inventory Alternates].AlternateID AS [Alternate ID],  [Inventory Alternates].Description AS Keterangan, [Inventory Alternates].UOM, [Inventory Alternates].QTY FROM         [Inventory Alternates] INNER JOIN                       Inventory ON [Inventory Alternates].NoItem = Inventory.NoItem WHERE     ([Inventory Alternates].NoItem = N'" & Param & "') and ([Inventory Alternates].TypeItem=0) ORDER BY [Inventory Alternates].AlternateID", CNN, lckLockBatch
Set MyDDE.ChildRecordset = RcAlter.DBRecordset.Clone(adLockBatchOptimistic)
Set DataGrid1(0).DataSource = MyDDE.ChildRecordset
End Sub

Private Sub OpenCompliment(ByVal Param As String)
RcCompli.DBOpen "SELECT     [Inventory Alternates].AlternateID AS [Alternate ID],  [Inventory Alternates].Description AS Keterangan, [Inventory Alternates].UOM, [Inventory Alternates].QTY FROM         [Inventory Alternates] INNER JOIN                       Inventory ON [Inventory Alternates].NoItem = Inventory.NoItem WHERE     ([Inventory Alternates].NoItem = N'" & Param & "')  and ([Inventory Alternates].TypeItem=1) ORDER BY [Inventory Alternates].AlternateID ", CNN, lckLockBatch
'Set MyDDE.ChildRecordset = RcCompli.DBRecordset.Clone(adLockBatchOptimistic)
'Set DataGrid1(1).DataSource = MyDDE.ChildRecordset
End Sub

Private Sub OpenCust(ByVal Param As String)
RcCust.DBOpen "SELECT     [Inventory Partner Alternates].PartnerID AS ID, PartnerDB.CompanyName AS [Nama Perusahaan], [Inventory Partner Alternates].RefCode AS [Ref Code],                       [Inventory Partner Alternates].[DESC] AS Keterangan, [Inventory Partner Alternates].UOM, [Inventory Partner Alternates].QTY FROM         [Inventory Partner Alternates] INNER JOIN                       PartnerDB ON [Inventory Partner Alternates].PartnerID = PartnerDB.PartnerID WHERE     ([Inventory Partner Alternates].TypePartner = 0) AND ([Inventory Partner Alternates].AlternateID = N'" & Param & "') ORDER BY [Inventory Partner Alternates].RefCode, [Inventory Partner Alternates].PartnerID", CNN, lckLockBatch
End Sub

Private Sub OpenPart(ByVal Param As String)
RcPart.DBOpen "SELECT     [Inventory Partner Alternates].PartnerID AS ID, PartnerDB.CompanyName AS [Nama Perusahaan], [Inventory Partner Alternates].RefCode AS [Ref Code],                       [Inventory Partner Alternates].[DESC] AS Keterangan, [Inventory Partner Alternates].UOM, [Inventory Partner Alternates].QTY FROM         [Inventory Partner Alternates] INNER JOIN                       PartnerDB ON [Inventory Partner Alternates].PartnerID = PartnerDB.PartnerID WHERE     ([Inventory Partner Alternates].TypePartner = 0) AND ([Inventory Partner Alternates].AlternateID = N'" & Param & "') ORDER BY [Inventory Partner Alternates].RefCode, [Inventory Partner Alternates].PartnerID", CNN, lckLockBatch
Set MyDDE.ChildRecordset = RcPart.DBRecordset.Clone(adLockBatchOptimistic)
Set DataGrid1(1).DataSource = MyDDE.ChildRecordset
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 1 Then Call SSTab2_Click(0)
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
Select Case SSTab2.Tab
       Case 0: Set MyDDE.ChildRecordset = RcAlter.DBRecordset.Clone(adLockBatchOptimistic)
       Case 1: Set MyDDE.ChildRecordset = RcPart.DBRecordset.Clone(adLockBatchOptimistic)
       Case 2: Set MyDDE.ChildRecordset = RcCust.DBRecordset.Clone(adLockBatchOptimistic)
       Case 3: Set MyDDE.ChildRecordset = RcPart.DBRecordset.Clone(adLockBatchOptimistic)
End Select
Set DataGrid1(SSTab2.Tab).DataSource = MyDDE.ChildRecordset
End Sub

Private Sub CariDataDefaultItem(ByVal Param As String)
Dim Rc As New DBQuick
Rc.DBOpen "SELECT     NoItem, ItemName, UOM FROM         Inventory WHERE     (PartnerID = N'" & Param & "')", CNN, lckLockReadOnly
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        MyDDE.ChildRecordset.Fields("Ref Code") = .Fields(0)
        MyDDE.ChildRecordset.Fields("Keterangan") = .Fields(1)
        MyDDE.ChildRecordset.Fields("uom") = .Fields(2)
        MyDDE.ChildRecordset.Fields("QTY") = 1
     End If
     .Close
End With
Set Rc = Nothing
End Sub

Private Sub GridLayout()
DataGrid1(0).Columns(0).width = 2055.118
DataGrid1(0).Columns(1).width = 3539.906
DataGrid1(0).Columns(2).width = 959.8111
DataGrid1(0).Columns(3).width = 1514.835
DataGrid1(1).Columns(0).width = 2445.166
DataGrid1(1).Columns(1).width = 1214.929
DataGrid1(1).Columns(2).width = 2129.953
DataGrid1(1).Columns(3).width = 975.1182
DataGrid1(1).Columns(4).width = 1289.764
End Sub
