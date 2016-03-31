VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmItemDescriptor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item Descriptor"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmItemDescriptor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   Tag             =   "Outsourced Reference"
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5355
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   5400
      Left            =   0
      ScaleHeight     =   5400
      ScaleWidth      =   9855
      TabIndex        =   9
      Top             =   0
      Width           =   9855
      Begin TabDlg.SSTab SSTab1 
         Height          =   5070
         Left            =   75
         TabIndex        =   10
         Top             =   105
         Width           =   9720
         _ExtentX        =   17145
         _ExtentY        =   8943
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "List"
         TabPicture(0)   =   "FrmItemDescriptor.frx":6852
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Picture3"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Detail"
         TabPicture(1)   =   "FrmItemDescriptor.frx":686E
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Picture4"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00EAAF6F&
            Height          =   4545
            Left            =   105
            ScaleHeight     =   4485
            ScaleWidth      =   9450
            TabIndex        =   13
            Top             =   405
            Width           =   9510
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "Desc ID"
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
               DataField       =   "Keterangan"
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
               DataField       =   "Tipe"
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
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "Unit Price"
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
               DataField       =   "Cost"
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
               DataField       =   "CurrID"
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
               Index           =   6
               Left            =   6645
               Locked          =   -1  'True
               MaxLength       =   13
               ScrollBars      =   2  'Vertical
               TabIndex        =   7
               Tag             =   "Partner"
               Text            =   "Text1"
               Top             =   1110
               Width           =   2250
            End
            Begin VB.CommandButton cmdLink 
               Enabled         =   0   'False
               Height          =   330
               Left            =   8565
               Picture         =   "FrmItemDescriptor.frx":688A
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   1065
               Visible         =   0   'False
               Width           =   330
            End
            Begin TabDlg.SSTab SSTab2 
               Height          =   2865
               Left            =   60
               TabIndex        =   8
               Top             =   1515
               Width           =   8925
               _ExtentX        =   15743
               _ExtentY        =   5054
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
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TabCaption(0)   =   "Alternate"
               TabPicture(0)   =   "FrmItemDescriptor.frx":6C14
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "Picture5"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).ControlCount=   1
               TabCaption(1)   =   "Complimentary"
               TabPicture(1)   =   "FrmItemDescriptor.frx":6C30
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Picture6"
               Tab(1).ControlCount=   1
               Begin VB.PictureBox Picture6 
                  Height          =   2400
                  Left            =   -74925
                  ScaleHeight     =   2340
                  ScaleWidth      =   8700
                  TabIndex        =   17
                  Top             =   375
                  Width           =   8760
                  Begin MSDataGridLib.DataGrid DataGrid1 
                     Bindings        =   "FrmItemDescriptor.frx":6C4C
                     Height          =   2370
                     Index           =   1
                     Left            =   0
                     TabIndex        =   18
                     Top             =   0
                     Width           =   8700
                     _ExtentX        =   15346
                     _ExtentY        =   4180
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
                     ColumnCount     =   7
                     BeginProperty Column00 
                        DataField       =   "Desc ID"
                        Caption         =   "Desc ID"
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
                        DataField       =   "Unit Price"
                        Caption         =   "Unit Price"
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
                        DataField       =   "Cost"
                        Caption         =   "Cost"
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
                        DataField       =   "TAX"
                        Caption         =   "TAX"
                        BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                           Type            =   5
                           Format          =   ""
                           HaveTrueFalseNull=   1
                           TrueValue       =   "YES"
                           FalseValue      =   "NO"
                           NullValue       =   "NO"
                           FirstDayOfWeek  =   0
                           FirstWeekOfYear =   0
                           LCID            =   1033
                           SubFormatType   =   7
                        EndProperty
                     EndProperty
                     BeginProperty Column05 
                        DataField       =   "Kode Supplier"
                        Caption         =   "Kode Supplier"
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
                        DataField       =   "Nama Perusahaan"
                        Caption         =   "Nama Perusahaan"
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
                           Alignment       =   1
                        EndProperty
                        BeginProperty Column03 
                           Alignment       =   1
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
               Begin VB.PictureBox Picture5 
                  Height          =   2400
                  Left            =   75
                  ScaleHeight     =   2340
                  ScaleWidth      =   8700
                  TabIndex        =   15
                  Top             =   375
                  Width           =   8760
                  Begin MSDataGridLib.DataGrid DataGrid1 
                     Bindings        =   "FrmItemDescriptor.frx":6C61
                     Height          =   2370
                     Index           =   0
                     Left            =   0
                     TabIndex        =   16
                     Top             =   0
                     Width           =   8700
                     _ExtentX        =   15346
                     _ExtentY        =   4180
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
                     ColumnCount     =   7
                     BeginProperty Column00 
                        DataField       =   "Desc ID"
                        Caption         =   "Desc ID"
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
                        DataField       =   "Unit Price"
                        Caption         =   "Unit Price"
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
                        DataField       =   "Cost"
                        Caption         =   "Cost"
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
                        DataField       =   "TAX"
                        Caption         =   "TAX"
                        BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                           Type            =   5
                           Format          =   ""
                           HaveTrueFalseNull=   1
                           TrueValue       =   "YES"
                           FalseValue      =   "NO"
                           NullValue       =   "NO"
                           FirstDayOfWeek  =   0
                           FirstWeekOfYear =   0
                           LCID            =   1033
                           SubFormatType   =   7
                        EndProperty
                     EndProperty
                     BeginProperty Column05 
                        DataField       =   "Kode Supplier"
                        Caption         =   "Kode Supplier"
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
                        DataField       =   "Nama Perusahaan"
                        Caption         =   "Nama Perusahaan"
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
                           Alignment       =   1
                        EndProperty
                        BeginProperty Column03 
                           Alignment       =   1
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
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Description"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   0
               Left            =   225
               TabIndex        =   25
               Top             =   495
               Width           =   795
            End
            Begin VB.Line Line1 
               Index           =   2
               X1              =   195
               X2              =   1620
               Y1              =   735
               Y2              =   735
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Assembly ID"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   1
               Left            =   225
               TabIndex        =   24
               Top             =   150
               Width           =   885
            End
            Begin VB.Line Line1 
               Index           =   0
               X1              =   195
               X2              =   1620
               Y1              =   390
               Y2              =   390
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Type"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   4
               Left            =   225
               TabIndex        =   23
               Top             =   840
               Width           =   360
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
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Unit Cost"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   2
               Left            =   225
               TabIndex        =   22
               Top             =   1185
               Width           =   660
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
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Unit Price"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   3
               Left            =   5505
               TabIndex        =   21
               Top             =   840
               Width           =   675
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "UOM"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   5
               Left            =   5505
               TabIndex        =   20
               Top             =   495
               Width           =   345
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Currency"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   6
               Left            =   5505
               TabIndex        =   19
               Top             =   1185
               Width           =   660
            End
            Begin VB.Line Line1 
               Index           =   6
               X1              =   5475
               X2              =   6900
               Y1              =   1425
               Y2              =   1425
            End
         End
         Begin VB.PictureBox Picture3 
            Height          =   4560
            Left            =   -74895
            ScaleHeight     =   4500
            ScaleWidth      =   9465
            TabIndex        =   11
            Top             =   405
            Width           =   9525
            Begin MSComctlLib.ListView ListView1 
               Height          =   4500
               Left            =   0
               TabIndex        =   12
               Top             =   0
               Width           =   9480
               _ExtentX        =   16722
               _ExtentY        =   7938
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
                  Text            =   "Desc ID"
                  Object.Width           =   2910
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Keterangan"
                  Object.Width           =   5644
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Tipe"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   3
                  Text            =   "UOM"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   4
                  Text            =   "Unit Price"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   5
                  Text            =   "Cost"
                  Object.Width           =   2540
               EndProperty
            End
         End
      End
   End
End
Attribute VB_Name = "FrmItemDescriptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsMytr                         As New DBQuick
Private RcAlter                         As New DBQuick
Private RcCompli                        As New DBQuick
Private RcPartner                       As New DBQuick
Private WithEvents mCall                As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private MyData                          As New clsTransaksi
Private MEdit As Boolean
Private mFirstCaller             As Boolean
Private mAccount                        As String

Private Sub cmdLink_Click()
OpenDetailPartner 1
End Sub

Private Sub DataGrid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub DataGrid1_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
If MEdit = True Then
   If DataGrid1(Index).col = 3 Then
      DataGrid1(Index).MarqueeStyle = dbgFloatingEditor
      DataGrid1(Index).AllowUpdate = True
   Else
      DataGrid1(Index).AllowUpdate = False
      DataGrid1(Index).MarqueeStyle = dbgHighlightRow
   End If
Else
   DataGrid1(Index).MarqueeStyle = dbgHighlightRow
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
'Call DGPurchase_KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_Load()
HiasFormManTell Picture2, Me
'HiasForm Picture1, Me
With MyDDE
     .EditModeReplace = False
     .SetPermissions = UserAddnewDenied
     Set .BindForm = FrmItemDescriptor
     .BindFormTAG = "Partner"
     Set .ActiveConnection = CNN
     .PrepareQuery = " SELECT     [Descriptor Header].DescID AS [Desc ID], [Descriptor Header].Description AS Keterangan, [Descriptor Header].TypeID AS Tipe,                        [Descriptor Header].UOM, [Descriptor Header].UnitPrice AS [Unit Price], [Descriptor Header].UnitCost AS Cost, [Descriptor Header].TAX,                        [Descriptor Header].PartnerID AS [Kode Supplier], PartnerDB.CompanyName AS [Nama Perusahaan],[Descriptor Header].CurrId FROM         [Descriptor Header] INNER JOIN                       PartnerDB ON [Descriptor Header].PartnerID = PartnerDB.PartnerID ORDER BY [Descriptor Header].DescID"
End With
SSTab1.Tab = 0
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
Set FrmItemDescriptor = Nothing
End Sub

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
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               PrepareQuery
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
                     If SendDataToServer(" Delete From  [Descriptor Alternates] WHERE     (DescID = N'" & txtBox(0) & "') AND (TypeDesc=0) ") = True Then
                     Do
                       If .EOF Then Exit Do
                          SendDataToServer " INSERT INTO [Descriptor Alternates]" & _
                                           " (AlterID, Description, PartnerID, UnitPrice, UnitCost, TAX, DescID, TypeDesc)" & _
                                           " VALUES  (N'" & .Fields("Desc ID") & "', N'" & .Fields("Keterangan") & "', N'" & .Fields("Kode Supplier") & "', " & CDbl(.Fields("Unit Price")) & ", " & CDbl(.Fields("Cost")) & ", " & BoolToInt(.Fields("TAX")) & ", N'" & txtBox(0) & "', 0)"
                     .MoveNext
                     Loop
                     End If
                     .MoveLast
                     End If
                End With
                
                With RcCompli.DBRecordset
                     If .Recordcount <> 0 Then
                     .MoveFirst
                     If SendDataToServer(" Delete From  [Descriptor Alternates] WHERE     (DescID = N'" & txtBox(0) & "') AND (TypeDesc=1) ") = True Then
                     Do
                       If .EOF Then Exit Do
                          SendDataToServer " INSERT INTO [Descriptor Alternates]" & _
                                           " (AlterID, Description, PartnerID, UnitPrice, UnitCost, TAX, DescID, TypeDesc)" & _
                                           " VALUES  (N'" & .Fields("Desc ID") & "', N'" & .Fields("Keterangan") & "', N'" & .Fields("Kode Supplier") & "', " & CDbl(.Fields("Unit Price")) & ", " & CDbl(.Fields("Cost")) & ", " & BoolToInt(.Fields("TAX")) & ", N'" & txtBox(0) & "', 1)"
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
               OpenDetailPartner 0
               MEdit = True
            End If
       Case tmbPrint:
            CallRPTReport "Descriptor Ref Table.rpt", "Select * From [Descriptor Ref Table] where [desc id] ='" & txtBox(0) & "'"
       Case tmbQuit:
            Unload Me
            Set MyDDE.BindForm = Nothing
End Select
txtBox(6).Enabled = False
cmdLink.Enabled = Not DataGrid1(0).Enabled
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenAlternate MyDDE.GetFieldByName("Desc ID")
OpenCompliment MyDDE.GetFieldByName("Desc ID")
End Sub

Private Sub OpenDetailPartner(ByVal Index As Integer)
On Error GoTo Hell:
Select Case Index
       Case 0:
            RcPartner.DBOpen "SELECT     [Descriptor Header].DescID AS [Desc ID], [Descriptor Header].Description AS Keterangan, [Descriptor Header].TypeID AS Tipe,                       [Descriptor Header].UOM, [Descriptor Header].UnitPrice AS [Unit Price], [Descriptor Header].UnitCost AS Cost, [Descriptor Header].TAX,                        [Descriptor Header].PartnerID AS [Kode Supplier], PartnerDB.CompanyName AS [Nama Perusahaan] FROM         [Descriptor Header] INNER JOIN                      PartnerDB ON [Descriptor Header].PartnerID = PartnerDB.PartnerID WHERE     ([Descriptor Header].DescID <> N'" & txtBox(0) & "') ORDER BY [Descriptor Header].DescID", CNN, lckLockReadOnly
            mFirstCaller = True
       Case 1: RcPartner.DBOpen " SELECT     CurrID, [Currency Name] FROM         [Currency Setup] ORDER BY CurrID", CNN, lckLockBatch
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 0: mCall.FromTagActive = "DESCRIPTOR ALTERNATE"
          Case 1: mCall.FromTagActive = "Master Currency"
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me
Else
   MessageBox "Data Belum Ada.", "Peringatan", msgOkOnly, msgCrtical
End If
Exit Sub
Hell:
'    messagebox Err.Description
    Err.Clear
End Sub

Private Sub PrepareQuery()
On Error Resume Next
With MyDDE

    .PrepareUpdate = " UPDATE [Descriptor Alternates] Set Currid =N'" & txtBox(6) & "', Description=N'" & txtBox(2) & "' WHERE     ([DescID] = N'" & txtBox(0) & "')"

    .PrepareDelete = " DELETE FROM [Descriptor Alternates] WHERE   ([DescID] = N'" & txtBox(0) & "') "
End With
End Sub

Private Sub ListView1_DblClick()
If MyDDE.ActiveRecordset.Recordcount <> 0 Then
   SSTab1.Tab = 1
   SSTab2.Tab = 0
End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
If MyDDE.ActiveRecordset.Recordcount <> 0 Then
   MyDDE.FindStringData "[Desc ID]='" & Item.Text & "'"
End If
End Sub

Private Sub mCall_BeforeUnload()
Select Case mCall.FromTagActive
       Case "DESCRIPTOR ALTERNATE":
'            If FindOwnRecordset(MyDDE.ChildRecordset, "[Desc ID] = '" & MyDDE.ChildRecordset.Fields("Desc ID") & "'") = True Then
'               MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Desc ID") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
'               MyDDE.ChildRecordset.CancelBatch adAffectCurrent
'               If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
'            End If

            If FindOwnRecordset(MyDDE.ChildRecordset, "[Desc ID] = '" & MyDDE.ChildRecordset.Fields("Desc ID") & "'") = True Then
               MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Desc ID") & " Sudah Ada....! Silahkan Diulangi", "Peringatan", msgOkOnly
               MyDDE.ChildRecordset.CancelBatch adAffectCurrent
               If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
            Else
               If IsNull(MyDDE.ChildRecordset.Fields("Desc ID")) = True Or MyDDE.ChildRecordset.Fields("Desc ID") = "" Then
                  'If MyDDE.ChildRecordset.Fields("Desc ID") = "" Then
                     MyDDE.ChildRecordset.CancelBatch adAffectCurrent
                     If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
                  'End If
               End If
            End If
            mFirstCaller = False
        Case "Master Currency":
            
End Select
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case mCall.FromTagActive
       Case "DESCRIPTOR ALTERNATE":
            MyDDE.ChildRecordset.Fields("Desc ID") = mCall.GetFieldByName(0)
            MyDDE.ChildRecordset.Fields("Keterangan") = mCall.GetFieldByName(1)
            MyDDE.ChildRecordset.Fields("Unit Price") = mCall.GetFieldByName("Unit Price")
            MyDDE.ChildRecordset.Fields("Cost") = mCall.GetFieldByName("Cost")
            MyDDE.ChildRecordset.Fields("Tax") = mCall.GetFieldByName("tax")
            MyDDE.ChildRecordset.Fields("Kode Supplier") = mCall.GetFieldByName("Kode Supplier")
            MyDDE.ChildRecordset.Fields("Nama Perusahaan") = mCall.GetFieldByName("Nama Perusahaan")
       Case "Master Currency":
            MyDDE.ChildRecordset.Fields("CurrID") = mCall.GetFieldByName(0)
            
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
                 .SubItems(3) = Avdata(3, I)
                 .SubItems(4) = FormatNumber(Avdata(4, I), 0)
                 .SubItems(5) = FormatNumber(Avdata(5, I), 0)
            End With
        Next I
     Else
     End If
End With
End Sub

Private Sub OpenAlternate(ByVal Param As String)
RcAlter.DBOpen "SELECT [Descriptor Alternates].AlterID AS [Desc ID], [Descriptor Alternates].Description AS Keterangan, [Descriptor Alternates].PartnerID AS [Kode Supplier],  PartnerDB.CompanyName AS [Nama Perusahaan], [Descriptor Alternates].UnitPrice AS [Unit Price], [Descriptor Alternates].UnitCost AS Cost,                        [Descriptor Alternates].TAX FROM         [Descriptor Alternates] INNER JOIN                       PartnerDB ON [Descriptor Alternates].PartnerID = PartnerDB.PartnerID WHERE     ([Descriptor Alternates].DescID = N'" & Param & "') AND ([Descriptor Alternates].TypeDesc = 0)  ORDER BY [Descriptor Alternates].AlterID", CNN, lckLockBatch
SSTab2_Click 0
End Sub

Private Sub OpenCompliment(ByVal Param As String)
RcCompli.DBOpen "SELECT [Descriptor Alternates].AlterID AS [Desc ID], [Descriptor Alternates].Description AS Keterangan, [Descriptor Alternates].PartnerID AS [Kode Supplier],  PartnerDB.CompanyName AS [Nama Perusahaan], [Descriptor Alternates].UnitPrice AS [Unit Price], [Descriptor Alternates].UnitCost AS Cost,                        [Descriptor Alternates].TAX FROM         [Descriptor Alternates] INNER JOIN                       PartnerDB ON [Descriptor Alternates].PartnerID = PartnerDB.PartnerID WHERE     ([Descriptor Alternates].DescID = N'" & Param & "') AND ([Descriptor Alternates].TypeDesc = 1) ORDER BY [Descriptor Alternates].AlterID", CNN, lckLockBatch
SSTab2_Click 0
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 1 Then Call SSTab2_Click(0)
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
If SSTab2.Tab = 0 Then
   Set MyDDE.ChildRecordset = RcAlter.DBRecordset.Clone(adLockBatchOptimistic)
   Set DataGrid1(0).DataSource = MyDDE.ChildRecordset
Else
   Set MyDDE.ChildRecordset = RcCompli.DBRecordset.Clone(adLockBatchOptimistic)
   Set DataGrid1(1).DataSource = MyDDE.ChildRecordset
End If
End Sub

Private Function OpenPartner(ByVal Index As Integer) As Boolean
Select Case Index
       Case 1: RcPartner.DBOpen " SELECT     CurrID, [Currency Name] FROM         [Currency Setup] ORDER BY CurrID", CNN, lckLockBatch
End Select
'MessageBox "SELECT WareHouse.WareHouse AS [Kode Gudang Persediaan], WareHouse.[WareHouse Name] AS [Nama Gudang Persediaan],                        WareHouse.NoAccount AS [Kode Perkiraan] FROM         [Inventory Group] INNER JOIN                       GLAccount ON [Inventory Group].NoAccount = GLAccount.NoAccount INNER JOIN                       WareHouse ON GLAccount.NoAccount = WareHouse.GroupAccount WHERE     ([Inventory Group].NoGroup = N'" & MyDDE.GetFieldByName("Kode Kelompok") & "')"
If RcPartner.Recordcount <> 0 Then
    Select Case Index
           Case 1: mCall.FromTagActive = "Master Currency"
    End Select
    Set mCall.FormData = RcPartner.DBRecordset
    mCall.LookUp Me
    'mLastGDG = ""
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgCrtical
   OpenPartner = True
End If
End Function
