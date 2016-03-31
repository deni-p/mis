VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmGELL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GELLIFICATION"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6225
      Left            =   0
      ScaleHeight     =   6225
      ScaleWidth      =   11640
      TabIndex        =   1
      Top             =   0
      Width           =   11640
      Begin VB.TextBox lblNoEkstraksi 
         Appearance      =   0  'Flat
         DataField       =   "no_ekstraksi"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2145
         TabIndex        =   39
         Tag             =   "Gen"
         Top             =   165
         Width           =   2055
      End
      Begin VB.TextBox txtkempu 
         Appearance      =   0  'Flat
         DataField       =   "total_kempu"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   1290
         TabIndex        =   36
         Tag             =   "Gen"
         Top             =   5760
         Width           =   2055
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         DataField       =   "Keterangan"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4560
         TabIndex        =   4
         Tag             =   "Gen"
         Top             =   5745
         Width           =   6990
      End
      Begin VB.TextBox txtGroup 
         Appearance      =   0  'Flat
         DataField       =   "grup"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2145
         TabIndex        =   3
         Tag             =   "Gen"
         Top             =   840
         Width           =   2055
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4080
         Left            =   120
         TabIndex        =   2
         Top             =   1575
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   7197
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         Tab             =   2
         TabsPerRow      =   5
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
         TabCaption(0)   =   "Bahan Penunjang untuk Proses Gellification"
         TabPicture(0)   =   "FrmGELL.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame1"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Plate Heat Exchanger (PHE)"
         TabPicture(1)   =   "FrmGELL.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame2"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Colling System"
         TabPicture(2)   =   "FrmGELL.frx":0038
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Frame3"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Gel Collector"
         TabPicture(3)   =   "FrmGELL.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame4"
         Tab(3).ControlCount=   1
         Begin VB.Frame Frame4 
            Height          =   3615
            Left            =   -74880
            TabIndex        =   31
            Top             =   360
            Width           =   11190
            Begin MSDataGridLib.DataGrid DataGrid4 
               Height          =   3225
               Left            =   135
               TabIndex        =   32
               Top             =   255
               Width           =   10920
               _ExtentX        =   19262
               _ExtentY        =   5689
               _Version        =   393216
               HeadLines       =   1
               RowHeight       =   15
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
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
                     LCID            =   1057
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
               EndProperty
            End
         End
         Begin VB.Frame Frame3 
            Height          =   3600
            Left            =   135
            TabIndex        =   25
            Top             =   360
            Width           =   11175
            Begin MSDataGridLib.DataGrid DataGrid3 
               Height          =   2535
               Left            =   105
               TabIndex        =   30
               Top             =   960
               Width           =   10935
               _ExtentX        =   19288
               _ExtentY        =   4471
               _Version        =   393216
               HeadLines       =   1
               RowHeight       =   15
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
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
                     LCID            =   1057
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
               EndProperty
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               DataField       =   "waktu_mulai_cooling"
               DataSource      =   "MyDDE"
               Height          =   330
               Index           =   2
               Left            =   1440
               TabIndex        =   28
               Top             =   195
               Width           =   4125
               _ExtentX        =   7276
               _ExtentY        =   582
               _Version        =   393216
               Format          =   63242242
               CurrentDate     =   39633
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               DataField       =   "waktu_selesai_cooling"
               DataSource      =   "MyDDE"
               Height          =   330
               Index           =   3
               Left            =   1440
               TabIndex        =   29
               Top             =   555
               Width           =   4125
               _ExtentX        =   7276
               _ExtentY        =   582
               _Version        =   393216
               Format          =   63242242
               CurrentDate     =   39633
            End
            Begin VB.Label lblReference 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Waktu Selesai"
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
               Height          =   195
               Index           =   7
               Left            =   225
               TabIndex        =   27
               Top             =   630
               Width           =   1005
            End
            Begin VB.Label lblReference 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Waktu Mulai"
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
               Height          =   195
               Index           =   6
               Left            =   210
               TabIndex        =   26
               Top             =   255
               Width           =   870
            End
         End
         Begin VB.Frame Frame2 
            Height          =   3645
            Left            =   -74895
            TabIndex        =   19
            Top             =   330
            Width           =   11235
            Begin MSDataGridLib.DataGrid DataGrid2 
               Height          =   2475
               Left            =   105
               TabIndex        =   24
               Top             =   1050
               Width           =   10995
               _ExtentX        =   19394
               _ExtentY        =   4366
               _Version        =   393216
               HeadLines       =   1
               RowHeight       =   15
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
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
                     LCID            =   1057
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
               EndProperty
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               DataField       =   "waktu_mulai_plate"
               DataSource      =   "MyDDE"
               Height          =   330
               Index           =   0
               Left            =   1470
               TabIndex        =   22
               Tag             =   "Gen"
               Top             =   225
               Width           =   4125
               _ExtentX        =   7276
               _ExtentY        =   582
               _Version        =   393216
               Format          =   63242242
               CurrentDate     =   39633
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               DataField       =   "waktu_selesai_plate"
               DataSource      =   "MyDDE"
               Height          =   330
               Index           =   1
               Left            =   1470
               TabIndex        =   23
               Top             =   585
               Width           =   4125
               _ExtentX        =   7276
               _ExtentY        =   582
               _Version        =   393216
               Format          =   63242242
               CurrentDate     =   39633
            End
            Begin VB.Label lblReference 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Waktu Selesai"
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
               Height          =   195
               Index           =   5
               Left            =   255
               TabIndex        =   21
               Top             =   660
               Width           =   1005
            End
            Begin VB.Label lblReference 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Waktu Mulai"
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
               Height          =   195
               Index           =   4
               Left            =   240
               TabIndex        =   20
               Top             =   285
               Width           =   870
            End
         End
         Begin VB.Frame Frame1 
            Height          =   3675
            Left            =   -74925
            TabIndex        =   14
            Top             =   315
            Width           =   11250
            Begin VB.TextBox txtOX 
               Appearance      =   0  'Flat
               DataField       =   "jml_ox05"
               DataSource      =   "MyDDE"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   1095
               TabIndex        =   35
               Tag             =   "Gen"
               Top             =   570
               Width           =   2055
            End
            Begin VB.TextBox txtair 
               Appearance      =   0  'Flat
               DataField       =   "jml_air"
               DataSource      =   "MyDDE"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   1095
               TabIndex        =   34
               Tag             =   "Gen"
               Top             =   240
               Width           =   2055
            End
            Begin MSDataGridLib.DataGrid dgDetail 
               Height          =   2475
               Left            =   60
               TabIndex        =   18
               Top             =   1140
               Width           =   11115
               _ExtentX        =   19606
               _ExtentY        =   4366
               _Version        =   393216
               HeadLines       =   1
               RowHeight       =   15
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
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
                     LCID            =   1057
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
               EndProperty
            End
            Begin VB.Label lblReference 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Cleaning Heat Tank (CAT)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   17
               Top             =   945
               Width           =   2160
            End
            Begin VB.Label lblReference 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "OX-05 (ml)"
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
               Height          =   195
               Index           =   2
               Left            =   240
               TabIndex        =   16
               Top             =   630
               Width           =   765
            End
            Begin VB.Label lblReference 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Air (Liter)"
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
               Height          =   195
               Index           =   1
               Left            =   315
               TabIndex        =   15
               Top             =   270
               Width           =   675
            End
         End
      End
      Begin MSComCtl2.DTPicker DcTanggal 
         DataField       =   "tanggal_press"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   2145
         TabIndex        =   5
         Tag             =   "Gen"
         Top             =   495
         Width           =   2070
         _ExtentX        =   3651
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
         Format          =   63242243
         CurrentDate     =   39634
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   7
         X1              =   9615
         X2              =   7485
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Label lblReference 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By"
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
         Height          =   195
         Index           =   9
         Left            =   7485
         TabIndex        =   38
         Top             =   240
         Width           =   930
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   9510
         TabIndex        =   37
         Top             =   180
         Width           =   2055
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   6
         X1              =   5685
         X2              =   3555
         Y1              =   6045
         Y2              =   6045
      End
      Begin VB.Label lblReference 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Kempu"
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
         Height          =   195
         Index           =   8
         Left            =   165
         TabIndex        =   33
         Top             =   5805
         Width           =   885
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   3
         X1              =   2295
         X2              =   165
         Y1              =   6060
         Y2              =   6060
      End
      Begin VB.Label lblKeterangan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
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
         Height          =   195
         Index           =   1
         Left            =   3600
         TabIndex        =   12
         Top             =   5790
         Width           =   840
      End
      Begin VB.Label lblGroup 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
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
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   885
         Width           =   435
      End
      Begin VB.Label lblTanggal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
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
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   540
         Width           =   570
      End
      Begin VB.Label lblid 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000005&
         Height          =   210
         Left            =   14640
         TabIndex        =   9
         Top             =   7560
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblFilterAid 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Ekstraksi"
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
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   210
         Width           =   930
      End
      Begin VB.Label LbRefID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "orderID"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2145
         TabIndex        =   7
         Tag             =   "Gen"
         Top             =   1170
         Width           =   2055
      End
      Begin VB.Label lblReference 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OrderID"
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
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   1230
         Width           =   585
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   0
         X1              =   2400
         X2              =   105
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   1
         X1              =   4185
         X2              =   120
         Y1              =   810
         Y2              =   810
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   2
         X1              =   2250
         X2              =   120
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   5
         X1              =   2250
         X2              =   120
         Y1              =   1485
         Y2              =   1485
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   6210
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   1005
      BindFormTAG     =   "FIL"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
   Begin VB.Label lblKeterangan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   585
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   1240
      X2              =   0
      Y1              =   255
      Y2              =   255
   End
End
Attribute VB_Name = "FrmGELL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim strSQL As String
Dim GridAltColor As String
Dim Changingsel As Byte
Dim Xval As String

Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1

Dim MEdit As Boolean
Dim sWCID As String

Private rsFind As New DBQuick
Private rsCleaning As New DBQuick
Private rsHeat As New DBQuick
Private rsCooling As New DBQuick
Private rsCollect As New DBQuick
Private rsPrecoating As New DBQuick


Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    ScanKey KeyCode, Shift, MyDDE

    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()

    HiasFormManTell Picture2, Me
    SSTab1.TabVisible(3) = False
    With MyDDE
        .EditModeReplace = False
        Set .BindForm = Me
        .BindFormTAG = "Gen"
        .SetPermissions = UserDeleteDenied
        Set .ActiveConnection = CNN

        .PrepareQuery = "SELECT * From Gellification"
        .SetPermissions = aksess.MayDo("Gellification")
    End With
    SSTab1.Tab = 0
End Sub

Private Sub PrepareQuery()
    On Error GoTo xErr
    Dim ket As Byte

    With MyDDE
      .PrepareAppend = "insert into GELLIFICATION(id_gell                 ,no_ekstraksi            , tgl_gell                                      ,grup                      ,              orderID     ,           jml_air             ,         jml_ox05             ,     waktu_mulai_cooling                                    ,          waktu_selesai_cooling                             ,                  waktu_mulai_plate                         ,                waktu_selesai_plate                         ,      keterangan             ,         total_kempu            ,issued_by) values ('" & _
                                                  lblNoEkstraksi.Text & "','" & lblNoEkstraksi & "','" & Format(DcTanggal.Value, "yyyy-MM-dd") & "','" & txtGroup(0).Text & "', '" & LbRefID.Caption & "', '" & FQty(txtair(1).Text) & "', '" & FQty(txtOX(2).Text) & "', '" & Format(DTPicker1(2).Value, "yyyy-MM-dd hh:mm:ss") & "', '" & Format(DTPicker1(3).Value, "yyyy-MM-dd hh:mm:ss") & "', '" & Format(DTPicker1(0).Value, "yyyy-MM-dd hh:mm:ss") & "', '" & Format(DTPicker1(1).Value, "yyyy-MM-dd hh:mm:ss") & "', '" & txtKeterangan.Text & "','" & FQty(txtkempu(3).Text) & "','" & MainMenu.StatusBar1.Panels(1).Text & "')"
                           
                           
      .PrepareUpdate = "update Gellification set tgl_gell = '" & Format(DcTanggal.Value, "yyyy-MM-dd") & "', grup='" & txtGroup(0).Text & "' orderID = '" & LbRefID.Caption & "' , jml_air = '" & FQty(txtair(1).Text) & "', jml_ox05 = '" & FQty(txtOX(2).Text) & "',  waktu_mulai_cooling = '" & Format(DTPicker1(2).Value, "yyyy-MM-dd hh:mm:ss") & "',waktu_selesai_cooling = '" & Format(DTPicker1(3).Value, "yyyy-MM-dd hh:mm:ss") & "' , waktu_mulai_plate = '" & Format(DTPicker1(0).Value, "yyyy-MM-dd hh:mm:ss") & "', waktu_selesai_plate = '" & Format(DTPicker1(1).Value, "yyyy-MM-dd hh:mm:ss") & "',   keterangan= '" & txtKeterangan.Text & "' , total_kempu =  '" & FQty(txtkempu(3).Text) & "'  where id_gell='" & lblNoEkstraksi & "'"
                           
      .PrepareDelete = "DELETE From Gellification Where id_gell = '" & lblNoEkstraksi & "'"
    End With

Exit Sub
xErr:
   Err.Clear
End Sub

Private Sub SimpanDetail()
   SendDataToServer "delete from gellification_cleaning_tank where id_gell='" & lblNoEkstraksi & "'"
   With rsCleaning.DBRecordset
      .MoveFirst
      While Not .EOF
         SendDataToServer "insert into gellification_cleaning_tank (id_gell,waktu,suhu) values ('" & _
                           lblNoEkstraksi & "', '" & Format(.Fields("waktu"), "yyyy-MM-dd hh:mm:ss") & "', '" & FQty(.Fields("suhu")) & "')"
         .MoveNext
      Wend
   End With
   
   SendDataToServer "delete from gellification_PITe where idGell ='" & lblNoEkstraksi & "'"
   With rsHeat.DBRecordset
      .MoveFirst
      While Not .EOF
          SendDataToServer "insert into gellification_PITe (idGell, waktu, suhu_masuk, suhu_keluar) values ('" & _
                           lblNoEkstraksi & "', '" & FQty(.Fields("waktu")) & "', '" & FQty(.Fields("suhu_masuk")) & "', '" & FQty(.Fields("suhu_keluar")) & "')"
         .MoveNext
      Wend
   End With
   
   SendDataToServer "delete from gellification_cooling where id_Gell='" & lblNoEkstraksi & "'"
   With rsCooling.DBRecordset
      .MoveFirst
      While Not .EOF
          SendDataToServer "insert into gellification_cooling (id_Gell, waktu, suhu_masuk, suhu_keluar, ph) values ('" & _
                           lblNoEkstraksi & "', '" & Format(.Fields("waktu"), "yyyy-MM-dd") & "', '" & FQty(.Fields("suhu_masuk")) & "', '" & FQty(.Fields("suhu_keluar")) & "', '" & FQty(.Fields("ph")) & "')"
         .MoveNext
      Wend
   End With
   
   SendDataToServer "delete from gellification_collector where idGell='" & lblNoEkstraksi & "'"
   With rsCollect.DBRecordset
      .MoveFirst
      While Not .EOF
          SendDataToServer "insert into gellification_collector (idGell, waktu, keras, lunak) values ('" & _
                           lblNoEkstraksi & "', '" & FQty(.Fields("waktu")) & "', '" & FQty(.Fields("keras")) & "', '" & FQty(.Fields("lunak")) & "')"
         .MoveNext
      Wend
   End With
   
End Sub


Private Function GetWC(ByVal FormIDNya As String)
    On Error GoTo Masjid
    Dim RcGetWC As New DBQuick
    RcGetWC.DBOpen "SELECT wcenter_header.WCID From wcenter_header Where wcenter_header.formid = 42", CNN, lckLockReadOnly
    sWCID = RcGetWC.DBRecordset.Fields("WCID")
    Exit Function
Masjid:
    MessageBox "Konfigurasi Gellification Kosong"
    Err.Clear
End Function

Private Sub OpenDetail(ByVal ParameterString As String)
    
    If ParameterString = "" Then ParameterString = "11111111111" ': Exit Sub
         rsCleaning.DBOpen "select * from gellification_cleaning_tank where id_gell='" & ParameterString & "'", CNN, lckLockBatch
         rsHeat.DBOpen "select * from gellification_PITe where idGell='" & ParameterString & "'", CNN, lckLockBatch
         rsCooling.DBOpen "select * from gellification_cooling where id_gell='" & ParameterString & "'", CNN, lckLockBatch
         rsCollect.DBOpen "select * from gellification_collector where idgell='" & ParameterString & "'", CNN, lckLockBatch
         
    
    Select Case SSTab1.Tab
      Case 0: Set MyDDE.ChildRecordset = rsCleaning.DBRecordset
      Case 1: Set MyDDE.ChildRecordset = rsHeat.DBRecordset
      Case 2: Set MyDDE.ChildRecordset = rsCooling.DBRecordset
      Case 3: Set MyDDE.ChildRecordset = rsCollect.DBRecordset

    End Select


    Set dgDetail.DataSource = MyDDE.ChildRecordset
    Set DataGrid2.DataSource = MyDDE.ChildRecordset
    Set DataGrid3.DataSource = MyDDE.ChildRecordset
    Set DataGrid4.DataSource = MyDDE.ChildRecordset
End Sub



Private Sub lblNoEkstraksi_LostFocus()
   Dim rsCek As New DBQuick
   rsCek.DBOpen "select * from statusProduksi where noEkstraksi='" & lblNoEkstraksi.Text & "'", CNN, lckLockBatch
   If rsCek.DBRecordset.Recordcount > 0 Then
      rsCek.DBOpen "select * from Gellification where no_Ekstraksi='" & lblNoEkstraksi.Text & "'", CNN, lckLockBatch
      If rsCek.DBRecordset.Recordcount > 0 Then
         MessageBox "Nomor Ekstraksi Ini Sudah Diinput...!", "Peringatan", msgOkOnly, msgCrtical
         lblNoEkstraksi.Text = ""
      End If
   Else
      MessageBox "Nomor Ekstraksi Ini tidak ditemukan...!", "Peringatan", msgOkOnly, msgCrtical
      lblNoEkstraksi.Text = ""
   End If
   rsCek.CloseDB
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
    Dim IDGen As New IDGenerator

    Select Case AdReasonActiveDb

        Case tmbSave
            SimpanDetail
            SaveToMO
            SSTab1.Tab = 0
        Case tmbAddNew
        
            rsFind.DBOpen "select * from GELLIFICATION where no_ekstraksi =  '" & lblNoEkstraksi & "'", CNN
            
            If rsFind.DBRecordset.EOF Then
            
            MEdit = True
            Me.Tag = "baru"
            LbRefID.Caption = frmProduksi.txtBox(5)
            'LblNoEkstraksi = frmProduksi.txtBox(1)
            MyDDE.GetFieldByName("id_gell") = lblNoEkstraksi
            MyDDE.GetFieldByName("keterangan") = "-"
            SetDataDetail
            
            Else
            MsgBox "No Ekstraksi ini sudah ada, silahkan Entry Ekstrkasi baru ", vbInformation, "Konfirmasi "
            Exit Sub
            End If

        Case tmbDelete
            PrepareQuery
    End Select

End Sub

Private Sub SetDataDetail()
   '*** detail
   rsCleaning.DBRecordset.AddNew
   rsCleaning.DBRecordset.Fields("id_gell") = lblNoEkstraksi
   rsCleaning.DBRecordset.Fields("waktu") = 30
   
   rsCleaning.DBRecordset.AddNew
   rsCleaning.DBRecordset.Fields("id_gell") = lblNoEkstraksi
   rsCleaning.DBRecordset.Fields("waktu") = 60
   
   rsCleaning.DBRecordset.AddNew
   rsCleaning.DBRecordset.Fields("id_gell") = lblNoEkstraksi
   rsCleaning.DBRecordset.Fields("waktu") = 90
   
   rsCleaning.DBRecordset.AddNew
   rsCleaning.DBRecordset.Fields("id_gell") = lblNoEkstraksi
   rsCleaning.DBRecordset.Fields("waktu") = 120
   
   rsCleaning.DBRecordset.AddNew
   rsCleaning.DBRecordset.Fields("id_gell") = lblNoEkstraksi
   rsCleaning.DBRecordset.Fields("waktu") = 150

   rsCleaning.DBRecordset.AddNew
   rsCleaning.DBRecordset.Fields("id_gell") = lblNoEkstraksi
   rsCleaning.DBRecordset.Fields("waktu") = 180
   

   
   '*** Plate heat
   rsHeat.DBRecordset.AddNew
   rsHeat.DBRecordset.Fields("idGell") = lblNoEkstraksi
   rsHeat.DBRecordset.Fields("waktu") = 0
   
   rsHeat.DBRecordset.AddNew
   rsHeat.DBRecordset.Fields("idGell") = lblNoEkstraksi
   rsHeat.DBRecordset.Fields("waktu") = 30
   
   rsHeat.DBRecordset.AddNew
   rsHeat.DBRecordset.Fields("idGell") = lblNoEkstraksi
   rsHeat.DBRecordset.Fields("waktu") = 60
   
   rsHeat.DBRecordset.AddNew
   rsHeat.DBRecordset.Fields("idGell") = lblNoEkstraksi
   rsHeat.DBRecordset.Fields("waktu") = 80
   
   rsHeat.DBRecordset.AddNew
   rsHeat.DBRecordset.Fields("idGell") = lblNoEkstraksi
   rsHeat.DBRecordset.Fields("waktu") = 120

   rsHeat.DBRecordset.AddNew
   rsHeat.DBRecordset.Fields("idGell") = lblNoEkstraksi
   rsHeat.DBRecordset.Fields("waktu") = 150
   
   rsHeat.DBRecordset.AddNew
   rsHeat.DBRecordset.Fields("idGell") = lblNoEkstraksi
   rsHeat.DBRecordset.Fields("waktu") = 180
   
   
   '*** Colling System
   rsCooling.DBRecordset.AddNew
   rsCooling.DBRecordset.Fields("id_gell") = lblNoEkstraksi
   rsCooling.DBRecordset.Fields("waktu") = 0
   
   rsCooling.DBRecordset.AddNew
   rsCooling.DBRecordset.Fields("id_gell") = lblNoEkstraksi
   rsCooling.DBRecordset.Fields("waktu") = 30
   
   rsCooling.DBRecordset.AddNew
   rsCooling.DBRecordset.Fields("id_gell") = lblNoEkstraksi
   rsCooling.DBRecordset.Fields("waktu") = 60
   
   rsCooling.DBRecordset.AddNew
   rsCooling.DBRecordset.Fields("id_gell") = lblNoEkstraksi
   rsCooling.DBRecordset.Fields("waktu") = 80
   
   rsCooling.DBRecordset.AddNew
   rsCooling.DBRecordset.Fields("id_gell") = lblNoEkstraksi
   rsCooling.DBRecordset.Fields("waktu") = 120

   rsCooling.DBRecordset.AddNew
   rsCooling.DBRecordset.Fields("id_gell") = lblNoEkstraksi
   rsCooling.DBRecordset.Fields("waktu") = 150
   
   rsCooling.DBRecordset.AddNew
   rsCooling.DBRecordset.Fields("id_gell") = lblNoEkstraksi
   rsCooling.DBRecordset.Fields("waktu") = 180
   
'*** Collector
   rsCollect.DBRecordset.AddNew
   rsCollect.DBRecordset.Fields("idGell") = lblNoEkstraksi
   rsCollect.DBRecordset.Fields("waktu") = "20 Menit Ke 1"
   
   rsCollect.DBRecordset.AddNew
   rsCollect.DBRecordset.Fields("idGell") = lblNoEkstraksi
   rsCollect.DBRecordset.Fields("waktu") = "20 Menit Ke 2"
   
   rsCollect.DBRecordset.AddNew
   rsCollect.DBRecordset.Fields("idGell") = lblNoEkstraksi
   rsCollect.DBRecordset.Fields("waktu") = "20 Menit Ke 3"
   
   rsCollect.DBRecordset.AddNew
   rsCollect.DBRecordset.Fields("idGell") = lblNoEkstraksi
   rsCollect.DBRecordset.Fields("waktu") = "20 Menit Ke 4"
   
   rsCollect.DBRecordset.AddNew
   rsCollect.DBRecordset.Fields("idGell") = lblNoEkstraksi
   rsCollect.DBRecordset.Fields("waktu") = "20 Menit Ke 5"

   rsCollect.DBRecordset.AddNew
   rsCollect.DBRecordset.Fields("idGell") = lblNoEkstraksi
   rsCollect.DBRecordset.Fields("waktu") = "20 Menit Ke 6"
   
   rsCollect.DBRecordset.AddNew
   rsCollect.DBRecordset.Fields("idGell") = lblNoEkstraksi
   rsCollect.DBRecordset.Fields("waktu") = "20 Menit Ke 7"
   
   rsCollect.DBRecordset.AddNew
   rsCollect.DBRecordset.Fields("idGell") = lblNoEkstraksi
   rsCollect.DBRecordset.Fields("waktu") = "20 Menit Ke 8"

   rsCollect.DBRecordset.AddNew
   rsCollect.DBRecordset.Fields("idGell") = lblNoEkstraksi
   rsCollect.DBRecordset.Fields("waktu") = "20 Menit Ke 9"
   
   rsCollect.DBRecordset.AddNew
   rsCollect.DBRecordset.Fields("idGell") = lblNoEkstraksi
   rsCollect.DBRecordset.Fields("waktu") = "20 Menit Ke 10"
End Sub

Private Sub SaveToMO()
    Dim dStart As Date
    Dim dFinish As Date
    Dim ActualTime As Double
    Dim rsCek As New DBQuick
    Dim sWCID As String
   
    ActualTime = Val(SelisihHariJam(DTPicker1(3), DTPicker1(0), 2))
   
    rsCek.DBOpen "select WCID from WCenter_Header where FormID = 42", CNN

    If rsCek.DBRecordset.Recordcount > 0 Then
        sWCID = rsCek.DBRecordset.Fields(0)
        SendDataToServer "update [order output detail] set actual_time=" & ActualTime & " where OrderID='" & LbRefID.Caption & "' and WCID='" & sWCID & "'"
    End If

    rsCek.CloseDB
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb

        Case tmbSave

            If MyDDE.CheckEmptyControl = False Then
                MyDDE.IsChildMemberReady = True
                SimpanDetail
                SaveToMO
            Else
                MyDDE.IsChildMemberReady = False
            End If

    End Select

End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
                               ByVal pError As ADODB.Error, _
                               adStatus As ADODB.EventStatusEnum, _
                               ByVal pRecordset As ADODB.Recordset)


    OpenDetail IIf(IsNull(MyDDE.GetFieldByName("id_gell")), "", MyDDE.GetFieldByName("id_gell"))
    Label1.Caption = IIf(IsNull(MyDDE.GetFieldByName("approved_by")), "", MyDDE.GetFieldByName("approved_by"))
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb

        Case tmbSave

            If MyDDE.CheckEmptyControl = False Then
                MyDDE.IsChildMemberReady = True
                PrepareQuery
                
            Else
                MyDDE.IsChildMemberReady = False
            End If

        Case tmbDelete
            PrepareQuery
    End Select

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   Select Case SSTab1.Tab
      Case 0: Set MyDDE.ChildRecordset = rsCleaning.DBRecordset
      Case 1: Set MyDDE.ChildRecordset = rsHeat.DBRecordset
      Case 2: Set MyDDE.ChildRecordset = rsCooling.DBRecordset
      Case 3: Set MyDDE.ChildRecordset = rsCollect.DBRecordset

   End Select

   Set dgDetail.DataSource = MyDDE.ChildRecordset
   Set DataGrid2.DataSource = MyDDE.ChildRecordset
   Set DataGrid3.DataSource = MyDDE.ChildRecordset
   Set DataGrid4.DataSource = MyDDE.ChildRecordset
End Sub



