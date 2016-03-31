VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmEvaluasiSupplier 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Evaluasi Supplier"
   ClientHeight    =   7950
   ClientLeft      =   1635
   ClientTop       =   1920
   ClientWidth     =   11985
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000015&
   Icon            =   "FrmEvaluasiSupplier.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   Tag             =   "Purchase Order"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7395
      Left            =   0
      ScaleHeight     =   7395
      ScaleWidth      =   11985
      TabIndex        =   43
      Top             =   0
      Width           =   11985
      Begin TabDlg.SSTab SSTab1 
         Height          =   7245
         Left            =   60
         TabIndex        =   21
         Top             =   60
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   12779
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   15380335
         TabCaption(0)   =   "     Data     "
         TabPicture(0)   =   "FrmEvaluasiSupplier.frx":6852
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Picture1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "     Setup     "
         TabPicture(1)   =   "FrmEvaluasiSupplier.frx":686E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Picture3"
         Tab(1).ControlCount=   1
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00EAAF6F&
            Height          =   6750
            Left            =   -74910
            ScaleHeight     =   6690
            ScaleWidth      =   11595
            TabIndex        =   56
            Top             =   390
            Width           =   11655
            Begin VB.CommandButton cmd 
               Caption         =   "Reset"
               Height          =   540
               Index           =   1
               Left            =   1425
               Picture         =   "FrmEvaluasiSupplier.frx":688A
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   5910
               Width           =   885
            End
            Begin VB.CommandButton cmd 
               Caption         =   "Simpan"
               Height          =   540
               Index           =   0
               Left            =   540
               Picture         =   "FrmEvaluasiSupplier.frx":D0DC
               TabIndex        =   19
               Top             =   5910
               Width           =   885
            End
            Begin VB.TextBox txtSetup 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   315
               Index           =   17
               Left            =   6435
               TabIndex        =   18
               Text            =   "0"
               Top             =   4920
               Width           =   765
            End
            Begin VB.TextBox txtSetup 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   315
               Index           =   16
               Left            =   4950
               TabIndex        =   17
               Text            =   "0"
               Top             =   4920
               Width           =   765
            End
            Begin VB.TextBox txtSetup 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   315
               Index           =   15
               Left            =   6435
               TabIndex        =   16
               Text            =   "0"
               Top             =   4545
               Width           =   765
            End
            Begin VB.TextBox txtSetup 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   315
               Index           =   14
               Left            =   4950
               TabIndex        =   15
               Text            =   "0"
               Top             =   4545
               Width           =   765
            End
            Begin VB.TextBox txtSetup 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   315
               Index           =   13
               Left            =   6435
               TabIndex        =   14
               Text            =   "0"
               Top             =   4155
               Width           =   765
            End
            Begin VB.TextBox txtSetup 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   315
               Index           =   12
               Left            =   4950
               TabIndex        =   13
               Text            =   "0"
               Top             =   4155
               Width           =   765
            End
            Begin VB.TextBox txtSetup 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   315
               Index           =   11
               Left            =   6465
               TabIndex        =   12
               Text            =   "0"
               Top             =   3165
               Width           =   765
            End
            Begin VB.TextBox txtSetup 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   315
               Index           =   10
               Left            =   4980
               TabIndex        =   11
               Text            =   "0"
               Top             =   3165
               Width           =   765
            End
            Begin VB.TextBox txtSetup 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   315
               Index           =   9
               Left            =   6465
               TabIndex        =   10
               Text            =   "0"
               Top             =   2790
               Width           =   765
            End
            Begin VB.TextBox txtSetup 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   315
               Index           =   8
               Left            =   4980
               TabIndex        =   9
               Text            =   "0"
               Top             =   2790
               Width           =   765
            End
            Begin VB.TextBox txtSetup 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   315
               Index           =   7
               Left            =   6465
               TabIndex        =   8
               Text            =   "0"
               Top             =   2400
               Width           =   765
            End
            Begin VB.TextBox txtSetup 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   315
               Index           =   6
               Left            =   4980
               TabIndex        =   7
               Text            =   "0"
               Top             =   2400
               Width           =   765
            End
            Begin VB.TextBox txtSetup 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   315
               Index           =   5
               Left            =   6480
               TabIndex        =   6
               Text            =   "0"
               Top             =   1335
               Width           =   765
            End
            Begin VB.TextBox txtSetup 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   4995
               TabIndex        =   5
               Text            =   "0"
               Top             =   1335
               Width           =   765
            End
            Begin VB.TextBox txtSetup 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   6480
               TabIndex        =   4
               Text            =   "0"
               Top             =   960
               Width           =   765
            End
            Begin VB.TextBox txtSetup 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   4995
               TabIndex        =   3
               Text            =   "0"
               Top             =   960
               Width           =   765
            End
            Begin VB.TextBox txtSetup 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   6480
               TabIndex        =   2
               Text            =   "0"
               Top             =   570
               Width           =   765
            End
            Begin VB.TextBox txtSetup 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   4995
               TabIndex        =   1
               Text            =   "0"
               Top             =   570
               Width           =   765
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "s/d"
               Height          =   360
               Index           =   20
               Left            =   5940
               TabIndex        =   77
               Top             =   4965
               Width           =   465
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "s/d"
               Height          =   360
               Index           =   19
               Left            =   5940
               TabIndex        =   76
               Top             =   4605
               Width           =   465
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "s/d"
               Height          =   360
               Index           =   18
               Left            =   5940
               TabIndex        =   75
               Top             =   4185
               Width           =   465
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Dianggap Buruk jika Selisih hari antara      "
               Height          =   360
               Index           =   17
               Left            =   750
               TabIndex        =   74
               Top             =   4965
               Width           =   6135
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Dianggap kurang baik jika Selisih hari antara       "
               Height          =   360
               Index           =   16
               Left            =   750
               TabIndex        =   73
               Top             =   4590
               Width           =   6570
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Dianggap baik jika Selisih hari antara          "
               Height          =   360
               Index           =   15
               Left            =   750
               TabIndex        =   72
               Top             =   4200
               Width           =   5985
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "s/d"
               Height          =   360
               Index           =   14
               Left            =   5970
               TabIndex        =   71
               Top             =   3210
               Width           =   465
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "s/d"
               Height          =   360
               Index           =   13
               Left            =   5970
               TabIndex        =   70
               Top             =   2850
               Width           =   465
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "s/d"
               Height          =   360
               Index           =   12
               Left            =   5970
               TabIndex        =   69
               Top             =   2430
               Width           =   465
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Dianggap Buruk jika prosentase Selisih antara      "
               Height          =   360
               Index           =   11
               Left            =   780
               TabIndex        =   68
               Top             =   3210
               Width           =   6135
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Dianggap kurang baik jika prosentase Selisih antara       "
               Height          =   360
               Index           =   10
               Left            =   780
               TabIndex        =   67
               Top             =   2835
               Width           =   6570
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Dianggap baik jika prosentase Selisih antara          "
               Height          =   360
               Index           =   2
               Left            =   780
               TabIndex        =   66
               Top             =   2445
               Width           =   5985
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "s/d"
               Height          =   360
               Index           =   9
               Left            =   5985
               TabIndex        =   65
               Top             =   1380
               Width           =   465
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "s/d"
               Height          =   360
               Index           =   8
               Left            =   5985
               TabIndex        =   64
               Top             =   1020
               Width           =   465
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "s/d"
               Height          =   360
               Index           =   1
               Left            =   5985
               TabIndex        =   63
               Top             =   600
               Width           =   465
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Dianggap Buruk jika prosentase penolakan antara      "
               Height          =   360
               Index           =   7
               Left            =   795
               TabIndex        =   62
               Top             =   1380
               Width           =   6135
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Dianggap kurang baik jika prosentase penolakan antara       "
               Height          =   360
               Index           =   6
               Left            =   795
               TabIndex        =   61
               Top             =   1005
               Width           =   6570
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Dianggap baik jika prosentase penolakan antara          "
               Height          =   360
               Index           =   5
               Left            =   795
               TabIndex        =   60
               Top             =   615
               Width           =   5985
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Kategori Ketepatan Waktu"
               Height          =   360
               Index           =   4
               Left            =   540
               TabIndex        =   59
               Top             =   3825
               Width           =   3510
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Kategori Ketepatan Ukuran"
               Height          =   360
               Index           =   3
               Left            =   540
               TabIndex        =   58
               Top             =   2085
               Width           =   3510
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Kategori Penolakan / Reject"
               Height          =   360
               Index           =   0
               Left            =   540
               TabIndex        =   57
               Top             =   285
               Width           =   3510
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00EAAF6F&
            Height          =   6735
            Left            =   60
            ScaleHeight     =   6675
            ScaleWidth      =   11625
            TabIndex        =   44
            Top             =   390
            Width           =   11685
            Begin VB.CommandButton Command1 
               Caption         =   "&Filter"
               Height          =   540
               Left            =   6075
               Picture         =   "FrmEvaluasiSupplier.frx":1392E
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   360
               Width           =   870
            End
            Begin VB.Frame Fr 
               BackColor       =   &H00EAAF6F&
               Caption         =   "Kategori Ketepatan Waktu"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   1215
               Index           =   4
               Left            =   8745
               TabIndex        =   49
               Top             =   5370
               Width           =   2655
               Begin VB.OptionButton op 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00EAAF6F&
                  Caption         =   "Kurang Baik"
                  ForeColor       =   &H80000008&
                  Height          =   300
                  Index           =   13
                  Left            =   120
                  TabIndex        =   41
                  Top             =   480
                  Width           =   1455
               End
               Begin VB.OptionButton op 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00EAAF6F&
                  Caption         =   "Baik"
                  ForeColor       =   &H80000008&
                  Height          =   300
                  Index           =   12
                  Left            =   120
                  TabIndex        =   40
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   1455
               End
               Begin VB.OptionButton op 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00EAAF6F&
                  Caption         =   "Buruk"
                  ForeColor       =   &H80000008&
                  Height          =   300
                  Index           =   14
                  Left            =   120
                  TabIndex        =   42
                  Top             =   705
                  Width           =   1455
               End
            End
            Begin VB.Frame Fr 
               BackColor       =   &H00EAAF6F&
               Caption         =   "Kategori Ketepatan"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   1215
               Index           =   3
               Left            =   6705
               TabIndex        =   48
               Top             =   5370
               Width           =   1935
               Begin VB.OptionButton op 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00EAAF6F&
                  Caption         =   "Buruk"
                  ForeColor       =   &H80000008&
                  Height          =   300
                  Index           =   11
                  Left            =   120
                  TabIndex        =   39
                  Top             =   750
                  Width           =   1455
               End
               Begin VB.OptionButton op 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00EAAF6F&
                  Caption         =   "Kurang Baik"
                  ForeColor       =   &H80000008&
                  Height          =   300
                  Index           =   10
                  Left            =   120
                  TabIndex        =   38
                  Top             =   495
                  Width           =   1455
               End
               Begin VB.OptionButton op 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00EAAF6F&
                  Caption         =   "Baik"
                  ForeColor       =   &H80000008&
                  Height          =   300
                  Index           =   9
                  Left            =   120
                  TabIndex        =   37
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   1455
               End
            End
            Begin VB.Frame Fr 
               BackColor       =   &H00EAAF6F&
               Caption         =   "Kategori Harga"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   1215
               Index           =   2
               Left            =   4785
               TabIndex        =   47
               Top             =   5370
               Width           =   1815
               Begin VB.OptionButton op 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00EAAF6F&
                  Caption         =   "Buruk"
                  ForeColor       =   &H80000008&
                  Height          =   300
                  Index           =   8
                  Left            =   120
                  TabIndex        =   36
                  Top             =   720
                  Width           =   1455
               End
               Begin VB.OptionButton op 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00EAAF6F&
                  Caption         =   "Baik"
                  ForeColor       =   &H80000008&
                  Height          =   300
                  Index           =   6
                  Left            =   120
                  TabIndex        =   34
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   1455
               End
               Begin VB.OptionButton op 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00EAAF6F&
                  Caption         =   "Kurang Baik"
                  ForeColor       =   &H80000008&
                  Height          =   300
                  Index           =   7
                  Left            =   120
                  TabIndex        =   35
                  Top             =   480
                  Width           =   1455
               End
            End
            Begin VB.Frame Fr 
               BackColor       =   &H00EAAF6F&
               Caption         =   "Kategori Komunikasi"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   1215
               Index           =   1
               Left            =   2745
               TabIndex        =   46
               Top             =   5370
               Width           =   1935
               Begin VB.OptionButton op 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00EAAF6F&
                  Caption         =   "Buruk"
                  ForeColor       =   &H80000008&
                  Height          =   300
                  Index           =   5
                  Left            =   120
                  TabIndex        =   33
                  Top             =   750
                  Width           =   1455
               End
               Begin VB.OptionButton op 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00EAAF6F&
                  Caption         =   "Kurang Baik"
                  ForeColor       =   &H80000008&
                  Height          =   300
                  Index           =   4
                  Left            =   120
                  TabIndex        =   32
                  Top             =   495
                  Width           =   1455
               End
               Begin VB.OptionButton op 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00EAAF6F&
                  Caption         =   "Baik"
                  ForeColor       =   &H80000008&
                  Height          =   300
                  Index           =   3
                  Left            =   120
                  TabIndex        =   31
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   1455
               End
            End
            Begin VB.Frame Fr 
               BackColor       =   &H00EAAF6F&
               Caption         =   "Kategori Penolakan / Reject"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   1215
               Index           =   0
               Left            =   225
               TabIndex        =   45
               Top             =   5370
               Width           =   2415
               Begin VB.OptionButton op 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00EAAF6F&
                  Caption         =   "Baik"
                  ForeColor       =   &H80000008&
                  Height          =   300
                  Index           =   0
                  Left            =   120
                  TabIndex        =   28
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   1455
               End
               Begin VB.OptionButton op 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00EAAF6F&
                  Caption         =   "Kurang Baik"
                  ForeColor       =   &H80000008&
                  Height          =   300
                  Index           =   1
                  Left            =   120
                  TabIndex        =   29
                  Top             =   495
                  Width           =   1455
               End
               Begin VB.OptionButton op 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00EAAF6F&
                  Caption         =   "Buruk"
                  ForeColor       =   &H80000008&
                  Height          =   300
                  Index           =   2
                  Left            =   120
                  TabIndex        =   30
                  Top             =   750
                  Width           =   1455
               End
            End
            Begin VB.CommandButton cmdLink 
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               Left            =   5040
               MaskColor       =   &H000000C0&
               Picture         =   "FrmEvaluasiSupplier.frx":1A180
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   143
               UseMaskColor    =   -1  'True
               Width           =   405
            End
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "CompanyName"
               DataSource      =   "MyDDE"
               Height          =   330
               Index           =   0
               Left            =   1725
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   22
               Tag             =   "SPP"
               Top             =   135
               Width           =   3315
            End
            Begin MSFlexGridLib.MSFlexGrid grdEvaluasi 
               Height          =   4245
               Left            =   210
               TabIndex        =   50
               Top             =   1020
               Width           =   11160
               _ExtentX        =   19685
               _ExtentY        =   7488
               _Version        =   393216
               Rows            =   3
               Cols            =   14
               FixedRows       =   2
               WordWrap        =   -1  'True
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               DataField       =   "period1"
               DataSource      =   "MyDDE"
               Height          =   330
               Left            =   1725
               TabIndex        =   24
               Tag             =   "SPP"
               Top             =   540
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   582
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
               CurrentDate     =   38272
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               DataField       =   "period2"
               DataSource      =   "MyDDE"
               Height          =   330
               Left            =   4065
               TabIndex        =   25
               Tag             =   "SPP"
               Top             =   540
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   582
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
               CurrentDate     =   38272
            End
            Begin MSComCtl2.DTPicker dtpAktual 
               DataField       =   "period1"
               DataSource      =   "MyDDE"
               Height          =   330
               Left            =   7530
               TabIndex        =   26
               Tag             =   "SPP"
               Top             =   510
               Visible         =   0   'False
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   582
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
               CurrentDate     =   38272
            End
            Begin VB.Label lblID 
               Caption         =   "ID"
               Height          =   255
               Left            =   7095
               TabIndex        =   55
               Top             =   120
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.Label lblPartner 
               Height          =   255
               Left            =   7170
               TabIndex        =   54
               Top             =   435
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "s/d"
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
               Index           =   2
               Left            =   3705
               TabIndex        =   53
               Top             =   615
               Width           =   255
            End
            Begin VB.Line Line1 
               Index           =   8
               X1              =   255
               X2              =   1755
               Y1              =   855
               Y2              =   855
            End
            Begin VB.Line Line1 
               Index           =   0
               X1              =   255
               X2              =   1755
               Y1              =   450
               Y2              =   450
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Periode"
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
               Left            =   255
               TabIndex        =   52
               Top             =   615
               Width           =   615
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nama Supplier"
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
               Left            =   225
               TabIndex        =   51
               Top             =   195
               Width           =   1155
            End
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   7380
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   1005
      BindFormTAG     =   "SPP"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmEvaluasiSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RcDetail                                          As New DBQuick
Private RcPartner                                         As New DBQuick
Private WithEvents mCall                                  As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private MyData                                            As New clsTransaksi
Private MEdit, mEditPO, mFirstCaller, mVarDetailPOClose   As Boolean
Private mAccount                                          As String
Private rsData                                            As New DBQuick
Private rsSetup                                           As New DBQuick

Dim NoUrut As Integer
Dim Xval As String

Private Sub loadSetup()
   Dim x As Integer
   rsSetup.DBOpen "SELECT  reject1, reject2, reject3, reject4, reject5, reject6, ukuran1, ukuran2, ukuran3, ukuran4, ukuran5, ukuran6, waktu1, waktu2, waktu3, waktu4, waktu5, " & _
                  "    waktu6 From setup_evaluasi_supplier ", CNN
   If rsSetup.DBRecordset.Recordcount > 0 Then
      For x = 0 To 17
         txtSetup(x).Text = rsSetup.DBRecordset.Fields(x)
      Next
   End If
End Sub


Private Sub headGrdEvaluasi()

    With grdEvaluasi
        .TextMatrix(0, 0) = "No"
        .TextMatrix(1, 0) = "No"
   
        .TextMatrix(0, 1) = "Nama Barang"
        .TextMatrix(1, 1) = "Nama Barang"
   
        .TextMatrix(0, 2) = "PO"
        .TextMatrix(1, 2) = "PO"
   
        .TextMatrix(0, 3) = "                                                 Jumlah"
        .TextMatrix(1, 3) = "Order"
   
        .TextMatrix(0, 4) = "                                                 Jumlah"
        .TextMatrix(1, 4) = "Datang"
   
        .TextMatrix(0, 5) = "                                                 Jumlah"
        .TextMatrix(1, 5) = "%Selisih"
   
        .TextMatrix(0, 6) = "                                                 Jumlah"
        .TextMatrix(1, 6) = "Reject"
   
        .TextMatrix(0, 7) = "                                                 Jumlah"
        .TextMatrix(1, 7) = "%Reject"
   
        .TextMatrix(0, 8) = "                      Tanggal"
        .TextMatrix(1, 8) = "Order"
   
        .TextMatrix(0, 9) = "                      Tanggal"
        .TextMatrix(1, 9) = "Rencana"
   
        .TextMatrix(0, 10) = "                      Tanggal"
        .TextMatrix(1, 10) = "Aktual"
   
        .TextMatrix(0, 11) = "Hari Keterlambatan"
        .TextMatrix(1, 11) = "Hari Keterlambatan"
   
        .TextMatrix(0, 12) = "Keterangan"
        .TextMatrix(1, 12) = "Keterangan"
   
        '.ColAlignment(-1) = flexAlignCenterCenter
      
        .MergeRow(0) = True
        .MergeCells = flexMergeRestrictRows
      
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(11) = True
        
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(12) = True
        
        .ColWidth(0) = 500
        .ColWidth(1) = 2500
        .ColWidth(2) = 2000
        .ColWidth(11) = 1500
        .ColWidth(12) = 2500
      
    End With

End Sub

Private Sub SimpanDetail()
    Dim nRow As Integer
On Error GoTo xErr
    With grdEvaluasi

        If SendDataToServer("DELETE FROM [evaluasi_supplier_detil] WHERE     (id = N'" _
                & lblid.Caption & "')") = True Then

            For nRow = 1 To grdEvaluasi.Rows - 2

                grdEvaluasi.row = nRow
                SendDataToServer _
                        " INSERT INTO [evaluasi_supplier_detil] (id, NoItem, NoPO, [jml_order],[jml_Datang],[jml_selisih],[jml_reject],[jml_selisih_reject],[tgl_order],[tgl_rencana],[tgl_aktual],[hari_terlambat],[keterangan]) " _
                        & " VALUES (N'" & lblid.Caption & "', N'" & _
                        grdEvaluasi.TextMatrix(nRow + 1, 13) & "', N'" & _
                        grdEvaluasi.TextMatrix(nRow + 1, 2) & "', N'" & _
                        FQty(grdEvaluasi.TextMatrix(nRow + 1, 3)) & "', N'" & _
                        FQty(grdEvaluasi.TextMatrix(nRow + 1, 4)) & "', N'" & _
                        FQty(grdEvaluasi.TextMatrix(nRow + 1, 5)) & "', N'" & _
                        FQty(grdEvaluasi.TextMatrix(nRow + 1, 6)) & "', N'" & _
                        FQty(grdEvaluasi.TextMatrix(nRow + 1, 7)) & "', N'" & _
                        grdEvaluasi.TextMatrix(nRow + 1, 8) & "', N'" & _
                        grdEvaluasi.TextMatrix(nRow + 1, 9) & "', N'" & _
                        grdEvaluasi.TextMatrix(nRow + 1, 10) & "', N'" & _
                        grdEvaluasi.TextMatrix(nRow + 1, 11) & "', N'" & _
                        grdEvaluasi.TextMatrix(nRow + 1, 12) & "')"
            Next nRow

        End If

    End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub LoadNewDataGrid(): Dim nRow  As Integer
    Dim RcDetail As New DBQuick
    Dim strFilter As String
    Dim rowCount As Integer
    Dim nReject As Double
    Dim nSelisih As Double
    Dim nWaktu As Double
    Dim x As Integer
On Error GoTo xErr
    strFilter = _
            "SELECT [PO Order].PurchaseID,Inventory.NoItem,Inventory.ItemName,[Detail PO].ScheduleDate, [Detail PO].QTYPO, [Detail PO].QTYReceive, [Detail PO].QTYRetur,  [PO Order].DatePurchase, " _
            & _
            "[PO Order].[Require Date], [Detail PO].receive_date From  [Detail PO]  INNER JOIN [PO Order] ON ([Detail PO].PurchaseID = [PO Order].PurchaseID)  INNER JOIN Inventory ON ([Detail PO].NoItem = Inventory.NoItem) Where " _
            & "[PO Order].PartnerID = '" & lblPartner.Caption & _
            "' AND [PO Order].DatePurchase Between '" & Format(DTPicker1, "yyyy-MM-dd") _
            & "' AND '" & Format(DTPicker2, "yyyy-MM-dd") & _
            "' AND [Detail PO].StatusTrans = '2'  Order By  Inventory.ItemName"
    RcDetail.DBOpen strFilter, CNN
    grdEvaluasi.Rows = 3
    
    With RcDetail.DBRecordset

        If .Recordcount > 0 Then
            grdEvaluasi.Rows = .Recordcount + 2
            .MoveFirst
            NoUrut = 1
            While Not .EOF
                grdEvaluasi.TextMatrix(NoUrut + 1, 0) = NoUrut                            ' No Urut
                grdEvaluasi.TextMatrix(NoUrut + 1, 1) = .Fields("itemName")               ' Nama Barang
                grdEvaluasi.TextMatrix(NoUrut + 1, 2) = .Fields("purchaseID")             ' PO No
                grdEvaluasi.TextMatrix(NoUrut + 1, 3) = .Fields("QtyPO")                  ' Qty PO
                grdEvaluasi.TextMatrix(NoUrut + 1, 4) = .Fields("QTYReceive")             ' Qty Terima
                grdEvaluasi.TextMatrix(NoUrut + 1, 5) = (Val(.Fields("QtyPO")) - Val(.Fields("QtyReceive"))) * 100 / _
                        Val(.Fields("QtyPO"))                                             ' Prosentase selisih
                grdEvaluasi.TextMatrix(NoUrut + 1, 6) = .Fields("QTYRetur")               ' Qty Reject
                grdEvaluasi.TextMatrix(NoUrut + 1, 7) = Val(.Fields("QtyRetur")) * 100 / _
                        Val(.Fields("QtyPO"))                                             ' prosentase Reject
                grdEvaluasi.TextMatrix(NoUrut + 1, 8) = Format(.Fields("datePurchase"), _
                        "dd MMM yy")                                                      ' Tgl Pesan
                grdEvaluasi.TextMatrix(NoUrut + 1, 9) = Format(.Fields("scheduleDate"), _
                        "dd MMM yy")                                                      ' Tgl Rencana datang
                grdEvaluasi.TextMatrix(NoUrut + 1, 10) = Format(.Fields( _
                        "Receive_date"), "dd MMM yy")                                     ' tgl brg diterima
                
                If .Fields("scheduleDate") >= .Fields("Receive_date") Then
                  grdEvaluasi.TextMatrix(NoUrut + 1, 11) = "0"
                Else
                  grdEvaluasi.TextMatrix(NoUrut + 1, 11) = SelisihHariJam(.Fields("scheduleDate"), .Fields("Receive_date"), 1)
                End If
                
                grdEvaluasi.TextMatrix(NoUrut + 1, 13) = .Fields("NoItem")
                NoUrut = NoUrut + 1
                .MoveNext
            Wend
            rowCount = grdEvaluasi.Rows - 1
            
            For x = 2 To rowCount
               nReject = nReject + Val(grdEvaluasi.TextMatrix(x, 7))
               nSelisih = nSelisih + Val(grdEvaluasi.TextMatrix(x, 5))
               nWaktu = nWaktu + Val(grdEvaluasi.TextMatrix(x, 11))
            Next
            
            nReject = nReject / rowCount
            nSelisih = nSelisih / rowCount
            nWaktu = nWaktu / rowCount
            
            'Kategori reject
            If (nReject >= Val(txtSetup(0).Text)) And (nReject < Val(txtSetup(1).Text)) Then
               Op(0).Value = True
               Op(1).Value = False
               Op(2).Value = False
            End If
            If (nReject >= Val(txtSetup(2).Text)) And (nReject < Val(txtSetup(3).Text)) Then
               Op(0).Value = False
               Op(1).Value = True
               Op(2).Value = False
            End If
            If (nReject >= Val(txtSetup(4).Text)) And (nReject < Val(txtSetup(5).Text)) Then
               Op(0).Value = False
               Op(1).Value = False
               Op(2).Value = True
            End If
            
            'kategori selisih
            If (nSelisih >= Val(txtSetup(6).Text)) And (nSelisih < Val(txtSetup(7).Text)) Then
               Op(9).Value = True
               Op(10).Value = False
               Op(11).Value = False
            End If
            If (nSelisih >= Val(txtSetup(8).Text)) And (nSelisih < Val(txtSetup(9).Text)) Then
               Op(9).Value = False
               Op(10).Value = True
               Op(11).Value = False
            End If
            If (nSelisih >= Val(txtSetup(10).Text)) And (nSelisih < Val(txtSetup(11).Text)) Then
               Op(9).Value = False
               Op(10).Value = False
               Op(11).Value = True
            End If
            
            'kategori ketepatan waktu
            If (nWaktu >= Val(txtSetup(12).Text)) And (nWaktu < Val(txtSetup(13).Text)) Then
               Op(12).Value = True
               Op(13).Value = False
               Op(14).Value = False
            End If
            If (nWaktu >= Val(txtSetup(14).Text)) And (nWaktu < Val(txtSetup(15).Text)) Then
               Op(12).Value = False
               Op(13).Value = True
               Op(14).Value = False
            End If
            If (nWaktu >= Val(txtSetup(16).Text)) And (nWaktu < Val(txtSetup(17).Text)) Then
               Op(12).Value = False
               Op(13).Value = False
               Op(14).Value = True
            End If
            
            
        Else

            For nRow = 1 To grdEvaluasi.Cols - 1
                grdEvaluasi.TextMatrix(2, nRow) = ""
            Next nRow

            MessageBox "Periode " & Format(DTPicker1, "dd-mm-yyyy") & " Hingga " & _
                    Format(DTPicker2, "dd-mm-yyyy") & " Transaksi Kosong..!", _
                    "Informasi", msgOkOnly, msgInfo
        End If

    End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub OpenDetail(ByVal ParameterString As String)
    Dim ncount As Integer
    Set RcDetail = New DBQuick
On Error GoTo xErr
    If ParameterString = "" Then ParameterString = "111" ': Exit Sub
    RcDetail.DBOpen _
            "SELECT evaluasi_supplier_detil.NoPO, evaluasi_supplier_detil.jml_order, evaluasi_supplier_detil.jml_Datang, evaluasi_supplier_detil.jml_selisih, evaluasi_supplier_detil.jml_reject,evaluasi_supplier_detil.jml_selisih_reject, evaluasi_supplier_detil.tgl_order, evaluasi_supplier_detil.tgl_rencana," _
            & _
            "evaluasi_supplier_detil.tgl_aktual,  evaluasi_supplier_detil.hari_terlambat, evaluasi_supplier_detil.keterangan, Inventory.ItemName From evaluasi_supplier_detil INNER JOIN evaluasi_supplier ON (evaluasi_supplier_detil.id = evaluasi_supplier.id) INNER JOIN Inventory ON (evaluasi_supplier_detil.NoItem = Inventory.NoItem) Where " _
            & "evaluasi_supplier.id = '" & ParameterString & _
            "' Order By evaluasi_supplier_detil.tgl_order", CNN, lckLockBatch
       
    Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
    grdEvaluasi.Rows = 3
   
    With RcDetail.DBRecordset

        If .Recordcount > 0 Then
            grdEvaluasi.Rows = .Recordcount + 2
            .MoveFirst
            NoUrut = 1
            While Not .EOF
                grdEvaluasi.TextMatrix(NoUrut + 1, 0) = NoUrut
                grdEvaluasi.TextMatrix(NoUrut + 1, 1) = .Fields("itemName")
                grdEvaluasi.TextMatrix(NoUrut + 1, 2) = .Fields("nopo")
                grdEvaluasi.TextMatrix(NoUrut + 1, 3) = .Fields("jml_order")
                grdEvaluasi.TextMatrix(NoUrut + 1, 4) = .Fields("jml_datang")
                grdEvaluasi.TextMatrix(NoUrut + 1, 5) = (.Fields("jml_datang") / _
                        .Fields("jml_order")) * 100
                grdEvaluasi.TextMatrix(NoUrut + 1, 6) = .Fields("jml_reject")
                grdEvaluasi.TextMatrix(NoUrut + 1, 7) = (.Fields("jml_reject") / _
                        .Fields("jml_order")) * 100
                grdEvaluasi.TextMatrix(NoUrut + 1, 8) = Format(.Fields("tgl_order"), _
                        "dd MMM yy")
                grdEvaluasi.TextMatrix(NoUrut + 1, 9) = Format(.Fields("tgl_rencana"), _
                        "dd MMM yy")
                grdEvaluasi.TextMatrix(NoUrut + 1, 10) = Format(.Fields("tgl_order"), _
                        "dd MMM yy")
                grdEvaluasi.TextMatrix(NoUrut + 1, 11) = Day(IIf(IsNull(.Fields( _
                        "tgl_order")), Now, .Fields("tgl_order")) - .Fields( _
                        "tgl_rencana"))
             '   grdEvaluasi.TextMatrix(NoUrut + 1, 12) = Day(IIf(IsNull(.Fields( _
                        "tgl_order")), Now, .Fields("tgl_order")) - .Fields( _
                        "tgl_rencana"))
                NoUrut = NoUrut + 1
                .MoveNext
            Wend
        End If

    End With

    grdEvaluasi.ColWidth(13) = 0
    RcDetail.CloseDB
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub LoadDataGrid()
    Dim NoUrut As Integer
    rsData.DBOpen "select * from view_evaluasi_detail where id=" & MyDDE.GetFieldByName( _
            "id"), CNN
   
    grdEvaluasi.Rows = 3
   
    With rsData.DBRecordset

        If .Recordcount > 0 Then
            .MoveFirst
            NoUrut = 1
            While Not .EOF
                grdEvaluasi.TextMatrix(0, NoUrut + 1) = NoUrut
                NoUrut = NoUrut + 1
                .MoveNext
            Wend
        End If

    End With

End Sub


Private Sub cmd_Click(Index As Integer)
On Error GoTo xErr
   If Index = 0 Then
      If SendDataToServer("update setup_evaluasi_supplier set reject1=" & txtSetup(0) & _
                                                      " ,reject2=" & txtSetup(1) & _
                                                      " ,reject3=" & txtSetup(2) & _
                                                      " ,reject4=" & txtSetup(3) & _
                                                      " ,reject5=" & txtSetup(4) & _
                                                      " ,reject6=" & txtSetup(5) & _
                                                      " ,ukuran1=" & txtSetup(6) & _
                                                      " ,ukuran2=" & txtSetup(7) & _
                                                      " ,ukuran3=" & txtSetup(8) & _
                                                      " ,ukuran4=" & txtSetup(9) & _
                                                      " ,ukuran5=" & txtSetup(10) & _
                                                      " ,ukuran6=" & txtSetup(11) & _
                                                      " ,waktu1 =" & txtSetup(12) & _
                                                      " ,waktu2 =" & txtSetup(13) & _
                                                      " ,waktu3 =" & txtSetup(14) & _
                                                      " ,waktu4 =" & txtSetup(15) & _
                                                      " ,waktu5 =" & txtSetup(16) & _
                                                      " ,waktu6 =" & txtSetup(17)) Then
            MessageBox "Data Telah tersimpan", "Informasi", msgOkOnly, msgInfo
         Else
            MessageBox "Data Gagal disimpan", "Error", msgOkOnly, msgCrtical
         End If
   Else
      loadSetup
   End If
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear
End Sub

Private Sub cmdLink_Click(Index As Integer)
    OpenPartner 0
End Sub

Private Sub Command1_Click()
    If lblPartner.Caption <> "" Then 'Trim(MyDDE.GetFieldByName("partnerID")) <> "" Then
        LoadNewDataGrid
    Else
        MessageBox "Pilih Supplier terlebih dahulu..!", "Peringatan", msgOkOnly, msgCrtical
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()

    HiasFormManTell Picture2, Me
    'HiasForm Picture1, Me

    headGrdEvaluasi
    
    loadSetup
    
    mVarDetailPOClose = False
    Set mCall = New frmCaller
    DTPicker1.Value = dDateBegin

    With MyDDE
        .EditModeReplace = False
        Set .BindForm = Me
        Set .ActiveConnection = CNN
        .PrepareQuery = " SELECT * from evaluasi_supplier "
        .SetPermissions = aksess.MayDo("Evaluasi Supplier")
    End With

    Set mCall = New frmCaller
    OpenDetail MyDDE.GetFieldByName("ID")
End Sub

Private Sub DetailSupplier(ByVal ParameterString As String)
    Set RcDetail = New DBQuick
    RcDetail.DBOpen _
            "SELECT PartnerDB.CompanyName From PartnerDB Where  PartnerDB.PartnerId=' " _
            & ParameterString & "'", CNN, lckLockBatch
    txtBox(0) = RcDetail.Fields("CompanyName")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Set MyData = Nothing
    MyDDE.ClearRecordset
    Set mCall = Nothing
End Sub

Private Sub grdEvaluasi_KeyPress(KeyAscii As Integer)
If grdEvaluasi.col = 12 And grdEvaluasi.Tag = "Edit" Then
      If KeyAscii = vbKeyReturn Then
         If grdEvaluasi.col + 1 = grdEvaluasi.Cols Then
            If grdEvaluasi.row + 1 = grdEvaluasi.Rows Then
               grdEvaluasi.row = 0
               grdEvaluasi.col = 0
            End If

            grdEvaluasi.row = grdEvaluasi.row + 1
            grdEvaluasi.col = 0
         Else
            grdEvaluasi.col = grdEvaluasi.col + 1
         End If
      End If
    
      If KeyAscii = 8 Then
         If Len(Xval) = 0 Then Exit Sub
         Xval = Left$(Xval, Len(Xval) - 1)
         Exit Sub
      End If

      Xval = Xval & Chr(KeyAscii)
   End If
End Sub

Private Sub grdEvaluasi_KeyUp(KeyCode As Integer, Shift As Integer)
If grdEvaluasi.col = 12 And grdEvaluasi.Tag = "Edit" Then
      grdEvaluasi.Text = Xval
   End If
End Sub

Private Sub grdEvaluasi_RowColChange()
Xval = ""
End Sub

Private Sub mCall_BeforeUnload()
    On Error Resume Next

    If FindOwnRecordset(MyDDE.ChildRecordset, "NoItem = '" & _
            MyDDE.ChildRecordset.Fields("NoItem") & "'") = True Then
        MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Noitem") & _
                " Sudah Ada....! Silahkan Diulangi", "Peringatan", msgOkOnly, msgCrtical
        MyDDE.ChildRecordset.CancelBatch adAffectCurrent

        If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
    Else

        If Not IsNull(MyDDE.ChildRecordset.Fields(0)) = True Then
            If MyDDE.ChildRecordset.Fields(0) = "" Then
                MyDDE.ChildRecordset.CancelBatch adAffectCurrent

                If MyDDE.ChildRecordset.Recordcount <> 0 Then _
                        MyDDE.ChildRecordset.MoveLast
            End If
        End If
    End If

    mFirstCaller = False
   
End Sub

Private Sub mCall_CallLinkForm()

    If mCall.FromTagActive = "Inventory List" Then
        FrmItemData.SetFocus
        FrmItemData.ZOrder (0)
    End If

End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, _
                               ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next

    If pRecordset.Recordcount <> 0 Then

        Select Case TagForm:
            
            Case "Supplier List":
                MyDDE.ChildRecordset.Fields("Partner ID") = mCall.GetFieldByName( _
                        "Partner ID")
                lblPartner.Caption = mCall.GetFieldByName("Partner ID")
                txtBox(0).Text = mCall.GetFieldByName("perusahaan")

            Case "Inventory List":
                MyDDE.ChildRecordset.Fields("SPPID") = MyDDE.GetFieldByName("SPPID")
                MyDDE.ChildRecordset.Fields("NoItem") = mCall.GetFieldByName("No barang")
                MyDDE.ChildRecordset.Fields("ItemName") = mCall.GetFieldByName( _
                        "Nama Barang")
                MyDDE.ChildRecordset.Fields("UOM") = mCall.GetFieldByName("UOM")
                MyDDE.ChildRecordset.Fields("Keperluan") = "-"
                MyDDE.ChildRecordset.Fields("Note") = "-"
                MyDDE.ChildRecordset.Fields("QTY_SPP") = 1
            
        End Select

    End If

End Sub

Private Sub MergeDoubleItem()
   
End Sub

Private Sub DGPurchase_KeyDown(KeyCode As Integer, _
                               Shift As Integer)

    If MEdit = False Then Exit Sub
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, _
                                    ByVal LastCol As Integer)

    If MEdit = False Then
    End If

End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
    PrepareQuery
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb

        Case tmbEdit, tmbDelete:

        Case tmbDetail:

            If txtBox(1) = "" Then MyDDE.GetFieldByName("Note") = "-"
            If MyDDE.CancelTrans = False Then
                If MyData.CheckGridKosong(MyDDE.ChildRecordset, "Qty_SPP") = True Then
                    MyDDE.CancelTrans = True
                    MessageBox "Data transaksi belum lengkap." & _
                            "Silahkan dicek kembali.", "Peringatan", msgOkOnly, msgCrtical
                End If
            End If

        Case tmbSave:

            If MyDDE.CheckEmptyControl = False Then
                If grdEvaluasi.Rows > 1 Then   'MyDDE.ChildRecordset.Recordcount <> 0 Then
                    MyDDE.IsChildMemberReady = True
                    PrepareQuery
                Else
                    MyDDE.IsChildMemberReady = False
                    MyDDE.CancelTrans = True
                End If

            Else
                MyDDE.IsChildMemberReady = False
            End If

    End Select

End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
    Dim nRow As Integer
    txtBox(0).Enabled = False

    Select Case AdReasonActiveDb

        Case tmbEdit:
            MEdit = True
            mEditPO = True
            
        Case tmbAddNew:
            MEdit = True
            lblid.Caption = IndexAuto
            SendDataToServer "delete from evaluasi_supplier"
            
            DTPicker1.Value = Now
            DTPicker2.Value = Now
            cmdLink(0).Enabled = True
            DTPicker1.SetFocus
            grdEvaluasi.Rows = 3
            For nRow = 1 To grdEvaluasi.Cols - 1
                grdEvaluasi.TextMatrix(2, nRow) = ""
            Next nRow
            grdEvaluasi.Tag = "Edit"
        Case tmbSave:

            If MyDDE.IsChildMemberReady = True Then
                SimpanDetail
                MEdit = False
                mEditPO = False
                OpenDetail txtBox(0)
                mVarDetailPOClose = False
                cmdLink(0).Enabled = False
            Else
                MessageBox "Detail Item  belum ada datanya.", "Peringatan", msgOkOnly, msgCrtical
            End If
            
        Case tmbCancel:

            If MyDDE.ChildRecordset.Recordcount = 0 Then
                MEdit = False
                mVarDetailPOClose = False
            End If

            cmdLink(0).Enabled = False
             
        Case tmbDetail:
            OpenPartner 3
            MEdit = True
            mVarDetailPOClose = False

        Case tmbPrint:
            Dim aReport As New utility
            aReport.CallReportView "select * from QueryEvaluasiSupplier where nopo='" & MyDDE.GetFieldByName("nopo") & "' and (tgl_order between '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' and '" & Format(DTPicker2.Value, "yyyy-MM-dd") & "')  ", "EvaluasiSupplier.rpt", ReportPath, "Evaluasi Supplier"
            Set aReport = Nothing
            'CallRPTReport "Purchase Order.rpt", "Select * From [purchase Order] where PurchaseID ='" & txtBox(0) & "'"
        Case tmbQuit:
            'Unload Me
            'Set MyDDE.BindForm = Nothing
        Case tmbCancel
           grdEvaluasi.Tag = ""
    End Select

    Err.Clear
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
                               ByVal pError As ADODB.Error, _
                               adStatus As ADODB.EventStatusEnum, _
                               ByVal pRecordset As ADODB.Recordset)
    OpenDetail MyDDE.GetFieldByName("ID")
    MEdit = False

    If MyDDE.ActiveRecordset.Recordcount > 0 Then
        Op(0).Value = IIf(MyDDE.GetFieldByName("reject") = 1, True, False)
        Op(1).Value = IIf(MyDDE.GetFieldByName("reject") = 2, True, False)
        Op(2).Value = IIf(MyDDE.GetFieldByName("reject") = 3, True, False)
        Op(3).Value = IIf(MyDDE.GetFieldByName("komunikasi") = 1, True, False)
        Op(4).Value = IIf(MyDDE.GetFieldByName("komunikasi") = 2, True, False)
        Op(5).Value = IIf(MyDDE.GetFieldByName("komunikasi") = 3, True, False)
        Op(6).Value = IIf(MyDDE.GetFieldByName("harga") = 1, True, False)
        Op(7).Value = IIf(MyDDE.GetFieldByName("harga") = 2, True, False)
        Op(8).Value = IIf(MyDDE.GetFieldByName("harga") = 3, True, False)
        Op(9).Value = IIf(MyDDE.GetFieldByName("tepat_volume") = 1, True, False)
        Op(10).Value = IIf(MyDDE.GetFieldByName("tepat_volume") = 2, True, False)
        Op(11).Value = IIf(MyDDE.GetFieldByName("tepat_volume") = 3, True, False)
        Op(12).Value = IIf(MyDDE.GetFieldByName("tepat_waktu") = 1, True, False)
        Op(13).Value = IIf(MyDDE.GetFieldByName("tepat_waktu") = 2, True, False)
        Op(14).Value = IIf(MyDDE.GetFieldByName("tepat_waktu") = 3, True, False)
        'LoadDataGrid
    End If

End Sub

Private Function IndexAuto() As String
   On Error Resume Next
   Dim Rc As New DBQuick
   Dim TglSaiki As String
   Dim Inom As String
   TglSaiki = Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2)
   Rc.DBOpen "SELECT  MAX(id) AS MaxNom FROM [evaluasi_supplier]", CNN, _
       lckLockReadOnly

   With Rc

      If .DBRecordset.Recordcount <> 0 Then
         Inom = IIf(Not IsNull(.Fields(0)), Mid(.DBRecordset.Fields("MaxNom"), 12, 5), "0") + 1

         If Err.Number = 94 Then Inom = 1
      Else
         Inom = 1
      End If

      Select Case Len(Trim(Str(Inom)))

         Case 0
            IndexAuto = "EVS-" & TglSaiki & "-" & Trim(Str(Inom))

         Case 1
            IndexAuto = "EVS-" & TglSaiki & "-" & "0000" & Trim(Str(Inom))

         Case 2
            IndexAuto = "EVS-" & TglSaiki & "-" & "000" & Trim(Str(Inom))

         Case 3
            IndexAuto = "EVS-" & TglSaiki & "-" & "00" & Trim(Str(Inom))

         Case 4
            IndexAuto = "EVS-" & TglSaiki & "-" & "0" & Trim(Str(Inom))
      End Select

   End With

End Function


Private Sub OpenPartner(ByVal Index As Integer)
    On Error GoTo Hell:

    Select Case Index

        Case 0:
            RcPartner.DBOpen MyData.UploadQuery("Supplier"), CNN, lckLockReadOnly

        Case 1:
            RcPartner.DBOpen MyData.UploadQuery("BANK", MyDDE.GetFieldByName( _
                    "PartnerID")), CNN, lckLockReadOnly

        Case 2:
            RcPartner.DBOpen _
                    "SELECT [Remainder PO].NoItem, Inventory.ItemName, Inventory.[Serial Supplier], [Remainder PO].QTYOrder, Inventory.PPn, Inventory.PriceIn * (Inventory.Markup / 100)   + Inventory.PriceIn AS Harga, [Remainder PO].SCNo FROM [Remainder PO] INNER JOIN Inventory ON [Remainder PO].NoItem = Inventory.NoItem ORDER BY [Remainder PO].NoItem", _
                    CNN, lckLockReadOnly

        Case 3:
            RcPartner.DBOpen _
                    "SELECT NoItem AS [No Barang], ItemName AS [Nama Barang], UOM, PPn,PriceIn AS Harga FROM         Inventory WHERE     (Manufacture = 0) ORDER BY NoItem", _
                    CNN, lckLockReadOnly

            'mFirstCaller = True
        Case 4:
            RcPartner.DBOpen _
                    "Select Code as Kode, Description as Keterangan,  [Bal_ Account Type], [Bal_ Account No_] from TermMethod ", _
                    CNN, lckLockReadOnly

        Case 5:
            RcPartner.DBOpen _
                    "Select No_ as Kode, Description as Keterangan, [Gen_ Prod_ Posting Group],  [Tax Group Code], [VAT Prod_ Posting Group], [Search Description], [Global Dimension 1 Code], [Global Dimension 2 Code] from item_charge ", _
                    CNN, lckLockReadOnly
    End Select

    If RcPartner.Recordcount <> 0 Then

        Select Case Index

            Case 0:
                mCall.FromTagActive = "Supplier List"
                mCall.CaptionLink = "Supplier"

            Case 1:
                mCall.FromTagActive = "Bank List"

            Case 2:
                mCall.FromTagActive = "Remindier"

            Case 3:
                mCall.FromTagActive = "Inventory List"
                mCall.CaptionLink = "Barang"

                'If MyDDE.ChildRecordset.Recordcount <> 0 Then mCall.txtCari = MyDDE.ChildRecordset.Fields("Noitem")
            Case 4:
                mCall.FromTagActive = "Term Method"
                mCall.CaptionLink = "Term Method"

            Case 5:
                mCall.FromTagActive = "Item Charge"
                mCall.CaptionLink = "Item Charge"
        End Select

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



Private Sub Picture1_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               Y As Single)
    'MoveForm Picture1.Parent.hwnd
End Sub

Function time_pos(ByVal obj_time As DTPicker, T As Integer)
    'obj_time.Move grdEvaluasi.Columns(T).Left + 100, (DGDETAIL.RowTop(DGDETAIL.Row) + DGDETAIL.Top)
    'obj_time.Move grdEvaluasi.Left
End Function

Private Sub PrepareQuery()

    On Error Resume Next
    Dim strSQL As String
    Dim katReject As Integer
    Dim katKomunikasi As Integer
    Dim katHarga As Integer
    Dim katKetepatan As Integer
    Dim katKetepatanWaktu As Integer

On Error GoTo xErr
    If Op(0).Value = True Then katReject = 1: If Op(1).Value = True Then katReject = 2: _
            If Op(2).Value = True Then katReject = 3
    
    If Op(3).Value = True Then katKomunikasi = 1: If Op(4).Value = True Then _
            katKomunikasi = 2: If Op(5).Value = True Then katKomunikasi = 3
    
    If Op(6).Value = True Then katHarga = 1: If Op(7).Value = True Then katHarga = 2: _
            If Op(8).Value = True Then katHarga = 3
    
    If Op(9).Value = True Then katKetepatan = 1: If Op(10).Value = True Then _
            katKetepatan = 2: If Op(11).Value = True Then katKetepatan = 3
    
    If Op(12).Value = True Then katKetepatanWaktu = 1: If Op(13).Value = True Then _
            katKetepatanWaktu = 2: If Op(14).Value = True Then katKetepatanWaktu = 3

    With MyDDE
        strSQL = _
                " INSERT INTO  [evaluasi_supplier] (id,period1,period2,supplier,reject,komunikasi,harga,[tepat_volume],[tepat_waktu],[CompanyName]) Values ('" _
                & lblid.Caption & "','" & Format(DTPicker1, "yyyy-MM-dd") & "','" & _
                Format(DTPicker2, "yyyy-MM-dd") & "','" & lblPartner & "','" & _
                katReject & "','" & katKomunikasi & "','" & katHarga & "','" & _
                katKetepatan & "','" & katKetepatanWaktu & "','" & txtBox(0) & "')"
        .PrepareAppend = strSQL
        
        strSQL = " UPDATE [evaluasi_supplier] set period1 = '" & Format(DTPicker1, _
                "yyyy-MM-dd") & "', period2 = '" & Format(DTPicker2, "yyyy-MM-dd") & _
                "', supplier = '" & lblPartner & "', reject = '" & katReject & _
                "', komunikasi = '" & katKomunikasi & "', harga = '" & katHarga & _
                "', [tepat_volume] = '" & katKetepatan & "', [tepat_waktu] = '" & _
                katKetepatanWaktu & "', [CompanyName] = '" & txtBox(0) & _
                "' where id ='" & lblid.Caption & "'"
        .PrepareUpdate = strSQL
                     
        .PrepareDelete = " DELETE FROM  [evaluasi_supplier] WHERE (id = '" & _
                lblid.Caption & "')"
    End With

    Err.Clear
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Function CekDetailItem(ByVal PoNumber As String, _
                               ByVal NoItemData As String) As Boolean
    Dim RcCek As New DBQuick
    RcCek.DBOpen "SELECT NoItem, SPPID FROM QuerySPP WHERE     (NoItem = N'" & _
            NoItemData & "') AND (SPPID = N'" & PoNumber & "')", CNN, lckLockReadOnly

    If RcCek.Recordcount <> 0 Then CekDetailItem = True
    RcCek.CloseDB
End Function

Private Function CekGridKosong() As Boolean
    Dim RcKsg As New DBQuick
    Dim Avdata As Variant
    Dim I As Integer
    Dim Temp As String
    Set RcKsg.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)

    With RcKsg

        If .Recordcount <> 0 Then
            Avdata = .DBRecordset.Getrows(.Recordcount, adBookmarkFirst)

            For I = 0 To UBound(Avdata, 2)
                Temp = IIf(Not IsNull(Avdata(0, I)), Avdata(0, I), "")

                If Temp <> "" Then
                    If Val(IIf(Not IsNull(Avdata(4, I)), Avdata(4, I), 0)) = 0 Then
                        MessageBox "Quantity harus diisi.", "Peringatan", msgOkOnly, msgCrtical
                        CekGridKosong = True
                        Exit For
                    End If

                Else
                    MessageBox "Data Item Tidak Lengkap.Harap Dicek Dulu", "Peringatan", msgOkOnly, msgCrtical
                    CekGridKosong = True
                    Exit For
                End If

            Next I

        Else
            CekGridKosong = True
        End If

    End With

    RcKsg.CloseDB
End Function

Private Function CekStock(ByVal NoItem As String) As Long
    Dim RcCek As New Recordset
    RcCek.CursorLocation = adUseClient
    RcCek.Open _
            "SELECT  SUM([Inventory Tabel].StockTmp)  AS QTY FROM [Inventory Tabel] INNER JOIN  Inventory ON [Inventory Tabel].NoItem = Inventory.NoItem GROUP BY [Inventory Tabel].NoItem, LEFT([Inventory Tabel].RefTrans, 2), Inventory.MinStock HAVING      (LEFT([Inventory Tabel].RefTrans, 2) = N'RN') AND ([Inventory Tabel].NoItem = N'" _
            & NoItem & "')", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText

    With RcCek

        If .Recordcount <> 0 Then
            CekStock = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
        Else
            CekStock = 0
        End If

        .Close
    End With

    Set RcCek = Nothing
End Function

