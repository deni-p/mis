VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmEnginering 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enginering"
   ClientHeight    =   6750
   ClientLeft      =   -15
   ClientTop       =   2310
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmEnginering.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   10455
   Tag             =   "Enginering Change Control"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6210
      Left            =   0
      ScaleHeight     =   6210
      ScaleWidth      =   10455
      TabIndex        =   14
      Top             =   0
      Width           =   10455
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
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
         Left            =   4380
         Picture         =   "FrmEnginering.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1170
         Width           =   330
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Reason"
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
         Height          =   720
         Index           =   2
         Left            =   1515
         MaxLength       =   249
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Tag             =   "Partner"
         Top             =   1530
         Width           =   6540
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "Methode"
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
         ItemData        =   "FrmEnginering.frx":6BDC
         Left            =   6765
         List            =   "FrmEnginering.frx":6BE6
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Tag             =   "Partner"
         Top             =   825
         Width           =   3225
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Directory"
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
         Left            =   1515
         Locked          =   -1  'True
         MaxLength       =   500
         MousePointer    =   2  'Cross
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Tag             =   "Partner"
         Top             =   1176
         Width           =   2865
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "BOM Change"
         DataField       =   "Bom Change"
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
         Height          =   315
         Index           =   1
         Left            =   1515
         TabIndex        =   3
         Tag             =   "Partner"
         Top             =   839
         Width           =   1635
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
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
         Left            =   9630
         Picture         =   "FrmEnginering.frx":6C08
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   150
         Width           =   330
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Status ECC"
         DataField       =   "StatusECC"
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
         Height          =   315
         Index           =   0
         Left            =   3180
         TabIndex        =   4
         Tag             =   "Partner"
         Top             =   839
         Width           =   1635
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Originator"
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
         Left            =   6765
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Tag             =   "Partner"
         Top             =   1185
         Width           =   3225
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "Efective Date"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1515
         TabIndex        =   2
         Tag             =   "Partner"
         Top             =   502
         Width           =   3300
         _ExtentX        =   5821
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
         Format          =   71630851
         CurrentDate     =   38560
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3600
         Left            =   120
         TabIndex        =   13
         Top             =   2370
         Width           =   10260
         _ExtentX        =   18098
         _ExtentY        =   6350
         _Version        =   393216
         Style           =   1
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
         TabCaption(0)   =   "Stage"
         TabPicture(0)   =   "FrmEnginering.frx":6F92
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Picture4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Component"
         TabPicture(1)   =   "FrmEnginering.frx":6FAE
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Picture5"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Approval"
         TabPicture(2)   =   "FrmEnginering.frx":6FCA
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Picture3"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.PictureBox Picture3 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3180
            Left            =   -74925
            ScaleHeight     =   3120
            ScaleWidth      =   10050
            TabIndex        =   23
            Top             =   360
            Width           =   10110
            Begin MSDataGridLib.DataGrid DataGrid1 
               Height          =   3090
               Index           =   2
               Left            =   0
               TabIndex        =   24
               Top             =   0
               Width           =   10020
               _ExtentX        =   17674
               _ExtentY        =   5450
               _Version        =   393216
               AllowUpdate     =   0   'False
               Appearance      =   0
               BorderStyle     =   0
               HeadLines       =   1
               RowHeight       =   15
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
               ColumnCount     =   3
               BeginProperty Column00 
                  DataField       =   "Departement"
                  Caption         =   "Departement"
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
                  DataField       =   "Employee"
                  Caption         =   "Approval By"
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
                  DataField       =   "Status"
                  Caption         =   "Status"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   5
                     Format          =   ""
                     HaveTrueFalseNull=   1
                     TrueValue       =   "Yes"
                     FalseValue      =   "No"
                     NullValue       =   "No"
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   7
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
                  BeginProperty Column02 
                     Alignment       =   1
                  EndProperty
               EndProperty
            End
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00EAAF6F&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3180
            Left            =   -74925
            ScaleHeight     =   3120
            ScaleWidth      =   10050
            TabIndex        =   19
            Top             =   360
            Width           =   10110
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "Keterangan"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   930
               Index           =   4
               Left            =   2685
               MaxLength       =   15
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   20
               Text            =   "FrmEnginering.frx":6FE6
               Top             =   3510
               Width           =   6360
            End
            Begin MSDataGridLib.DataGrid DataGrid1 
               Height          =   3120
               Index           =   1
               Left            =   0
               TabIndex        =   21
               Top             =   0
               Width           =   10050
               _ExtentX        =   17727
               _ExtentY        =   5503
               _Version        =   393216
               AllowUpdate     =   -1  'True
               Appearance      =   0
               BorderStyle     =   0
               HeadLines       =   1
               RowHeight       =   15
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
               ColumnCount     =   5
               BeginProperty Column00 
                  DataField       =   "SeqStageID"
                  Caption         =   "Stage"
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
                  DataField       =   "Komponen ID"
                  Caption         =   "Komponen ID"
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
                  DataField       =   "Nama Komponen"
                  Caption         =   "Nama Komponen"
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
               BeginProperty Column04 
                  DataField       =   "QTYUsage"
                  Caption         =   "QTY"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "#,##0;(#,##0)"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
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
                     Alignment       =   1
                  EndProperty
               EndProperty
            End
            Begin VB.Line Line1 
               Index           =   6
               X1              =   1365
               X2              =   2790
               Y1              =   4425
               Y2              =   4425
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Catatan"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   6
               Left            =   1365
               TabIndex        =   22
               Top             =   4170
               Width           =   630
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00EAAF6F&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3180
            Left            =   75
            ScaleHeight     =   3120
            ScaleWidth      =   10050
            TabIndex        =   15
            Top             =   360
            Width           =   10110
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "Catatan"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   930
               Index           =   3
               Left            =   2685
               MaxLength       =   200
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   16
               Text            =   "FrmEnginering.frx":6FEC
               Top             =   3510
               Width           =   6360
            End
            Begin MSDataGridLib.DataGrid DataGrid1 
               Bindings        =   "FrmEnginering.frx":6FF2
               Height          =   3120
               Index           =   0
               Left            =   0
               TabIndex        =   17
               Top             =   0
               Width           =   10050
               _ExtentX        =   17727
               _ExtentY        =   5503
               _Version        =   393216
               AllowUpdate     =   -1  'True
               Appearance      =   0
               BorderStyle     =   0
               HeadLines       =   1
               RowHeight       =   15
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
               ColumnCount     =   4
               BeginProperty Column00 
                  DataField       =   "NoLine"
                  Caption         =   "Seq. No"
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
                  DataField       =   "SeqStageID"
                  Caption         =   "Stage"
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
               BeginProperty Column03 
                  DataField       =   "ResourcesID"
                  Caption         =   "Resources"
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
                     ColumnWidth     =   4320
                  EndProperty
                  BeginProperty Column03 
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Catatan"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   8
               Left            =   1365
               TabIndex        =   18
               Top             =   4170
               Width           =   630
            End
            Begin VB.Line Line1 
               Index           =   8
               X1              =   1365
               X2              =   2790
               Y1              =   4425
               Y2              =   4425
            End
         End
      End
      Begin MSComDlg.CommonDialog Dialog 
         Left            =   9030
         Top             =   1590
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Document"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   7
         Left            =   135
         TabIndex        =   32
         Top             =   1236
         Width           =   945
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   135
         X2              =   1635
         Y1              =   1485
         Y2              =   1485
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   5385
         X2              =   6885
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Method"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   5
         Left            =   5385
         TabIndex        =   31
         Top             =   885
         Width           =   720
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   135
         X2              =   1635
         Y1              =   2235
         Y2              =   2235
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   5385
         X2              =   6885
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   5385
         X2              =   6885
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   135
         X2              =   1635
         Y1              =   802
         Y2              =   802
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reason"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   4
         Left            =   135
         TabIndex        =   30
         Top             =   1950
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Originator"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   3
         Left            =   5385
         TabIndex        =   29
         Top             =   1245
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F/G Item No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   2
         Left            =   5385
         TabIndex        =   28
         Top             =   210
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Efective Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   1
         Left            =   135
         TabIndex        =   27
         Top             =   554
         Width           =   1215
      End
      Begin VB.Label LblEcc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "NoItem"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   6765
         TabIndex        =   8
         Tag             =   "Partner"
         Top             =   150
         Width           =   2865
      End
      Begin VB.Label LblEcc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "ECCNo"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   1515
         TabIndex        =   1
         Tag             =   "Partner"
         Top             =   150
         Width           =   3300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ECC No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   26
         Top             =   210
         Width           =   645
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   135
         X2              =   1635
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Label LblEcc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "Description"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   2
         Left            =   6765
         TabIndex        =   10
         Tag             =   "Partner"
         Top             =   495
         Width           =   3225
      End
      Begin VB.Line Line1 
         Index           =   9
         X1              =   5370
         X2              =   6870
         Y1              =   810
         Y2              =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   9
         Left            =   5385
         TabIndex        =   25
         Top             =   540
         Width           =   1035
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   6180
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmEnginering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Rc As New DBQuick
Private RcComponent As New DBQuick
Private RcAproval As New DBQuick
Private RcPartner As New DBQuick
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private mKeyLoad As Boolean

Private Sub cmdLink_Click(Index As Integer)
On Error GoTo RepERR
If Index = 0 Then
   OpenPartner Index
ElseIf Index = 1 Then
    With dialog
        .InitDir = App.Path
        .Filter = "*.Doc|*.Doc" '"Crystal Report"
        .flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNPathMustExist
        .ShowOpen
        If .Filename = "" Then
        Else
            txtBox(1) = .Filename
        End If
    End With
End If
RepERR:
    If Err <> 0 Then
        MessageBox Err.Description & " - " & Err.Number, "Peringatan", msgOkOnly, msgCrtical
    End If
End Sub

Private Sub DataGrid1_ButtonClick(Index As Integer, ByVal ColIndex As Integer)
If DTPicker1.Enabled = True Then
Select Case Index
       Case 2:
            Select Case ColIndex
                   Case 0: OpenPartner 2
                   Case 1: OpenPartner 3
                   Case 2:
                        If DataGrid1(Index).Columns(ColIndex).Value = True Then
                           DataGrid1(Index).Columns(ColIndex).Value = False
                        Else
                           DataGrid1(Index).Columns(ColIndex).Value = True
                        End If
            End Select
End Select
End If
End Sub

Private Sub DataGrid1_Error(Index As Integer, ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub

Private Sub DataGrid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If mKeyLoad = True Then
   ScanKey KeyCode, Shift, MyDDE
Else
   Call Form_KeyDown(KeyCode, Shift)
End If
End Sub

Private Sub Form_Initialize()
Set mCall = New frmCaller
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If mKeyLoad = False Then mKeyLoad = True Else mKeyLoad = False
If mKeyLoad = False Then ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
SSTab1.Tab = 0
GridLayout
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
Check1(0).BackColor = &HEAAF6F
Check1(1).BackColor = &HEAAF6F
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmEnginering
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT     [Enginering Change].ECCNo, [Enginering Change].StatusECC, [Enginering Change].NoItem, Inventory.ItemName AS Description,   [Enginering Change].[Bom Change], [Enginering Change].[Efective Date], [Enginering Change].Originator, [Enginering Change].Directory,  [Enginering Change].Methode, [Enginering Change].Reason FROM         [Enginering Change] INNER JOIN Inventory ON [Enginering Change].NoItem = Inventory.NoItem"
End With
End Sub

Private Sub mCall_BeforeUnload()
On Error GoTo 1
Select Case mCall.FromTagActive
       Case "Departement":
            If FindOwnRecordset(MyDDE.ChildRecordset, "Departement = '" & MyDDE.ChildRecordset.Fields("Departement") & "'") = True Then
               MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Departement") & " Sudah Ada....! Silahkan Diulangi", "Peringatan", msgOkOnly, msgCrtical
               MyDDE.ChildRecordset.CancelBatch adAffectCurrent
               If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
            Else
               If IsNull(MyDDE.ChildRecordset.Fields(0)) = True Or MyDDE.ChildRecordset.Fields("Departement") = "" Then
'                  If MyDDE.ChildRecordset.Fields("Departement") = "" Then
                     MyDDE.ChildRecordset.CancelBatch adAffectCurrent
                     If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
'                  End If
               End If
            End If
End Select
Exit Sub
1:
MessageBox Err.Description, "frmengineering_mcall_beforeunload" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case TagForm
       Case "Bill Of Material":
            MyDDE.GetFieldByName("NoItem") = mCall.GetFieldByName(0)
            OpenDetail IIf(Not IsNull(mCall.GetFieldByName(0)), mCall.GetFieldByName(0), "xxxx")
            OpenDetailComponent IIf(Not IsNull(mCall.GetFieldByName(0)), mCall.GetFieldByName(0), "xxxx")
       Case "Departement":
            MyDDE.ChildRecordset.Fields("Kode Dep") = mCall.GetFieldByName("Dept ID")
            MyDDE.ChildRecordset.Fields("Departement") = mCall.GetFieldByName("Departement")
       Case "Karyawan":
            MyDDE.ChildRecordset.Fields("EmpID") = mCall.GetFieldByName("Kode Karyawan")
            MyDDE.ChildRecordset.Fields("Employee") = mCall.GetFieldByName("Nama Karyawan")
            
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
cmdLink(0).Enabled = DTPicker1.Enabled
cmdLink(1).Enabled = DTPicker1.Enabled
DataGrid1(2).Columns(0).Button = DTPicker1.Enabled
DataGrid1(2).Columns(1).Button = DTPicker1.Enabled
DataGrid1(2).Columns(2).Button = DTPicker1.Enabled
Select Case AdReasonActiveDb
       Case tmbEdit:
            SSTab1.Tab = 2
            cmdLink(0).SetFocus
       Case tmbAddNew:
            SSTab1.Tab = 2
            LblEcc(0).Caption = IndexAuto
            cmdLink(0).SetFocus
       Case tmbDetail:
            MyDDE.ChildRecordset.Fields("Status") = False
            OpenPartner 2
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
            With MyDDE.ChildRecordset
                If .Recordcount <> 0 Then
                    If SendDataToServer("DELETE FROM [Enginering Dept] WHERE     (ECCNo = N'" & LblEcc(0) & "')") = True Then
                        .MoveFirst
                        Do
                          If .EOF Then Exit Do
                           SendDataToServer ("INSERT INTO [Enginering Dept] (ECCNo, EmpID, [Kode Dep], Status) VALUES (N'" & LblEcc(0) & "', N'" & MyDDE.ChildRecordset.Fields("EmpID") & "', " & MyDDE.ChildRecordset.Fields("Kode Dep") & ", " & BoolToInt(MyDDE.ChildRecordset.Fields("Status")) & ")")
                          .MoveNext
                        Loop
                        .MoveLast
                    End If
                End If
            End With
            End If
End Select

End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbAddNew:
            SSTab1.Tab = 2
       Case tmbDetail:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               SSTab1.Tab = 2
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

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo xErr
MyDDE.PrepareAppend = " INSERT INTO [Enginering Change]" & _
                      " (ECCNo, StatusECC, NoItem, [Bom Change], [Efective Date], Originator, Directory, Methode, Reason)" & _
                      " VALUES (N'" & LblEcc(0) & "', " & BoolToInt(Check1(0)) & ", N'" & LblEcc(1) & "', " & BoolToInt(Check1(1)) & ", CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), N'" & txtBox(0) & "', N'" & txtBox(1) & "', N'" & Combo1 & "', N'" & txtBox(2) & "')"
                    
MyDDE.PrepareUpdate = " UPDATE [Enginering Change]" & _
                      " SET StatusECC = " & BoolToInt(Check1(0)) & ", NoItem = N'" & LblEcc(1) & "', [Bom Change] = " & BoolToInt(Check1(1)) & ", [Efective Date] = CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), Originator = N'" & txtBox(0) & "', Directory = N'" & txtBox(1) & "'," & _
                      " Methode = N'" & Combo1 & "', Reason = N'" & txtBox(2) & "'" & _
                      " WHERE (ECCNo = N'" & LblEcc(0) & "')"
                    
MyDDE.PrepareDelete = "DELETE FROM [Enginering Change] WHERE     (ECCNo = N'" & LblEcc(0) & "')"
Err.Clear
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If SSTab1.Tab = 0 Then
   OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("NoItem")), MyDDE.GetFieldByName("NoItem"), "xxxx")
Else
   OpenDetailComponent IIf(Not IsNull(MyDDE.GetFieldByName("NoItem")), MyDDE.GetFieldByName("NoItem"), "xxxx")
End If
Aproval
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

Private Function IndexAuto() As String
On Error GoTo 1
Dim Rc As New DBQuick
Dim TglSaiki As String
Dim Inom As Long
TglSaiki = Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2)
Rc.DBOpen " SELECT MAX(RIGHT(ECCNo, 5)) AS MaxNom FROM  [Enginering Change] WHERE (LEFT(ECCNo, 3) = N'ECC')", CNN, lckLockReadOnly
With Rc
     If .DBRecordset.Recordcount <> 0 Then
        Inom = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + 1
     Else
        Inom = 1
     End If
     Select Case Len(Trim(Str(Inom)))
            Case 0: IndexAuto = "ECC/" & TglSaiki & "-" & Trim(Str(Inom))
            Case 1: IndexAuto = "ECC/" & TglSaiki & "-" & "0000" & Trim(Str(Inom))
            Case 2: IndexAuto = "ECC/" & TglSaiki & "-" & "000" & Trim(Str(Inom))
            Case 3: IndexAuto = "ECC/" & TglSaiki & "-" & "00" & Trim(Str(Inom))
            Case 4: IndexAuto = "ECC/" & TglSaiki & "-" & "0" & Trim(Str(Inom))
     End Select
End With
Exit Function
1:
MessageBox Err.Description, "frmemployess:indexauto" & Err.Number, msgOkOnly, msgExclamation
End Function

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:
Select Case Index
       Case 0:
            RcPartner.DBOpen " SELECT NoItem AS [Kode Barang], ItemName AS [Nama Barang], Merk, UOM FROM         Inventory WHERE     (Manufacture = 1) ORDER BY NoItem", CNN, lckLockReadOnly
       Case 1:
            RcPartner.DBOpen " SELECT     NoItem AS [No Barang], ItemName AS [Nama Barang], UOM, PPn,PriceIn AS Harga FROM         Inventory WHERE     (Manufacture = 0) ORDER BY NoItem", CNN, lckLockReadOnly
       Case 2:
            RcPartner.DBOpen "SELECT     [Nama Dep] AS Departement, [Kode Dep] AS [Dept ID] FROM         [Tabel Departemen] ORDER BY [Nama Dep]", CNN, lckLockReadOnly
       Case 3:
            RcPartner.DBOpen "SELECT     FullName AS [Nama Karyawan], EmpID AS [Kode Karyawan] FROM         Employees WHERE     ([Kode dep] = " & MyDDE.ChildRecordset.Fields("Kode Dep") & ") ORDER BY FullName", CNN, lckLockReadOnly
'            mFirstCaller = True
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 0: mCall.FromTagActive = "Bill Of Material"
          Case 1:
          Case 2: mCall.FromTagActive = "Departement"
          Case 3: mCall.FromTagActive = "Karyawan"
   End Select
'   If MyDDE.ChildRecordset.Recordcount <> 0 Then mCall.txtCari = MyDDE.ChildRecordset.Fields("Noitem")
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me

Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgCrtical
   If MyDDE.ChildRecordset.Recordcount <> 0 Then
      MyDDE.ChildRecordset.CancelBatch adAffectCurrent
      If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
   End If
End If
'
Exit Sub
Hell:
    Err.Clear
End Sub

Private Sub OpenDetail(ByVal Param As String)
Rc.DBOpen "SELECT     [BOM Stage Detail].NoLine, [BOM Stage Detail].WCID as SeqStageID, [BOM Stage Detail].Description AS Keterangan, [BOM Stage Detail].ResourcesID,                       [Resources Table].Description AS Resources, [BOM Stage Detail].StageNote AS Catatan FROM         [BOM Stage Detail] INNER JOIN                      Inventory ON [BOM Stage Detail].NoItem = Inventory.NoItem AND [BOM Stage Detail].BomReff = Inventory.BomReff LEFT OUTER JOIN                       [Resources Table] ON [BOM Stage Detail].ResourcesID = [Resources Table].ResourcesID WHERE     (Inventory.NoItem = N'" & Param & "') ORDER BY [BOM Stage Detail].NoLine", CNN, lckLockBatch
Set DataGrid1(0).DataSource = Rc.DBRecordset
End Sub

Private Sub OpenDetailComponent(ByVal Param As String)
RcComponent.DBOpen "SELECT     [BOM Component Detail].SeqStageID, [BOM Component Detail].Description AS Keterangan, [BOM Component Detail].Component AS [Komponen ID],  Inventory.ItemName AS [Nama Komponen], [BOM Component Detail].UOM, [BOM Component Detail].QTYUsage, [BOM Component Detail].NoItem FROM         [BOM Component Detail] INNER JOIN Inventory ON [BOM Component Detail].Component = Inventory.NoItem AND [BOM Component Detail].BomReff = Inventory.BomReff LEFT OUTER JOIN [BOM Stage Detail] ON [BOM Component Detail].WCID = [BOM Stage Detail].WCID AND [BOM Component Detail].NoItem = [BOM Stage Detail].NoItem WHERE ([BOM Component Detail].NoItem = N'" & Param & "') ORDER BY [BOM Stage Detail].NoLine", CNN, lckLockBatch
Set DataGrid1(1).DataSource = RcComponent.DBRecordset
End Sub

Private Sub Aproval()
RcAproval.DBOpen "SELECT  [Tabel Departemen].[Nama Dep] as Departement , Employees.FullName as Employee,  [Enginering Dept].Status, [Enginering Dept].EmpID, [Enginering Dept].[Kode Dep],  [Enginering Dept].ECCNo FROM         [Enginering Dept] INNER JOIN  Employees ON [Enginering Dept].EmpID = Employees.EmpID INNER JOIN [Tabel Departemen] ON Employees.[Kode Dep] = [Tabel Departemen].[Kode Dep] WHERE     ([Enginering Dept].ECCNo = N'" & MyDDE.GetFieldByName("ECCNo") & "')", CNN, lckLockBatch
Set MyDDE.ChildRecordset = RcAproval.DBRecordset.Clone(adLockBatchOptimistic)
Set DataGrid1(2).DataSource = MyDDE.ChildRecordset
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
       Case 0: OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("NoItem")), MyDDE.GetFieldByName("NoItem"), "xxxx")
       Case 1: OpenDetailComponent IIf(Not IsNull(MyDDE.GetFieldByName("NoItem")), MyDDE.GetFieldByName("NoItem"), "xxxx")
       Case 2: Aproval
End Select
End Sub

Private Sub GridLayout()
DataGrid1(0).Columns(0).width = 810.1418
DataGrid1(0).Columns(1).width = 2145.26
DataGrid1(0).Columns(2).width = 4320
DataGrid1(0).Columns(3).width = 2174.74
DataGrid1(1).Columns(0).width = 1755.213
DataGrid1(1).Columns(1).width = 2204.788
DataGrid1(1).Columns(2).width = 3360.189
DataGrid1(1).Columns(3).width = 1019.906
DataGrid1(1).Columns(4).width = 1124.787
DataGrid1(2).Columns(0).width = 2415.118
DataGrid1(2).Columns(1).width = 5190.236
DataGrid1(2).Columns(2).width = 1844.787

End Sub

