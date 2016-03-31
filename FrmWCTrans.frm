VERSION 5.00
Object = "{FB992564-9055-42B5-B433-FEA84CEA93C4}#11.0#0"; "crviewer.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmWorkCenter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Work Center  Transaction"
   ClientHeight    =   5835
   ClientLeft      =   21495
   ClientTop       =   2280
   ClientWidth     =   10965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmWCTrans.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   Tag             =   "Work Centers"
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   19
      Top             =   5265
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   1005
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   5250
      Left            =   0
      ScaleHeight     =   5250
      ScaleWidth      =   10965
      TabIndex        =   29
      Top             =   0
      Width           =   10965
      Begin VB.TextBox txtBox 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DataField       =   "description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   9360
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Tag             =   "Partner"
         Text            =   "Text1"
         Top             =   120
         Width           =   2385
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5070
         Left            =   120
         TabIndex        =   0
         Top             =   105
         Width           =   10770
         _ExtentX        =   18997
         _ExtentY        =   8943
         _Version        =   393216
         Style           =   1
         Tabs            =   8
         TabsPerRow      =   9
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
         TabCaption(0)   =   "List Work Center"
         TabPicture(0)   =   "FrmWCTrans.frx":6852
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Picture5"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Detail"
         TabPicture(1)   =   "FrmWCTrans.frx":686E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Picture3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Shift"
         TabPicture(2)   =   "FrmWCTrans.frx":688A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Picture6"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Routings"
         TabPicture(3)   =   "FrmWCTrans.frx":68A6
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Picture4(1)"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Resources"
         TabPicture(4)   =   "FrmWCTrans.frx":68C2
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Picture4(0)"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "BOM Where Used"
         TabPicture(5)   =   "FrmWCTrans.frx":68DE
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "GrdBOMWhere"
         Tab(5).Control(0).Enabled=   0   'False
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "MO Where Used"
         TabPicture(6)   =   "FrmWCTrans.frx":68FA
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "GrdMOUsed"
         Tab(6).Control(0).Enabled=   0   'False
         Tab(6).ControlCount=   1
         TabCaption(7)   =   "Inquiry"
         TabPicture(7)   =   "FrmWCTrans.frx":6916
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "lblChartType"
         Tab(7).Control(0).Enabled=   0   'False
         Tab(7).Control(1)=   "CrWCInQuiry"
         Tab(7).Control(1).Enabled=   0   'False
         Tab(7).Control(2)=   "OptStackChart"
         Tab(7).Control(2).Enabled=   0   'False
         Tab(7).Control(3)=   "OptGanttChart"
         Tab(7).Control(3).Enabled=   0   'False
         Tab(7).Control(4)=   "pctReport"
         Tab(7).Control(4).Enabled=   0   'False
         Tab(7).ControlCount=   5
         Begin VB.PictureBox pctReport 
            BorderStyle     =   0  'None
            Height          =   4590
            Left            =   -75000
            ScaleHeight     =   4590
            ScaleWidth      =   9405
            TabIndex        =   79
            Top             =   360
            Width           =   9410
         End
         Begin VB.OptionButton OptGanttChart 
            Caption         =   "Gantt Chart"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -65520
            TabIndex        =   27
            Tag             =   "ana"
            Top             =   600
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton OptStackChart 
            Caption         =   "Stack Chart"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -65520
            TabIndex        =   28
            Tag             =   "ana"
            Top             =   870
            Width           =   1215
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00EAAF6F&
            Height          =   4545
            Left            =   -74900
            ScaleHeight     =   4485
            ScaleWidth      =   10500
            TabIndex        =   49
            Top             =   400
            Width           =   10560
            Begin VB.CommandButton cmdLink 
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               Left            =   8790
               MaskColor       =   &H000000C0&
               Picture         =   "FrmWCTrans.frx":6932
               Style           =   1  'Graphical
               TabIndex        =   6
               Tag             =   "SPPH"
               Top             =   450
               UseMaskColor    =   -1  'True
               Width           =   360
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00EAAF6F&
               Caption         =   " Movement "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2850
               Left            =   210
               TabIndex        =   59
               Top             =   1410
               Width           =   4620
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  DataField       =   "queue_time"
                  Height          =   315
                  Index           =   3
                  Left            =   2250
                  MaxLength       =   50
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   10
                  Tag             =   "Partner"
                  Top             =   1305
                  Width           =   1140
               End
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  DataField       =   "Utilization"
                  Height          =   315
                  Index           =   4
                  Left            =   2250
                  MaxLength       =   50
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   12
                  Tag             =   "Partner"
                  Top             =   2325
                  Width           =   1140
               End
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  DataField       =   "setup_time"
                  Height          =   315
                  Index           =   23
                  Left            =   2250
                  MaxLength       =   50
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   9
                  Tag             =   "Partner"
                  Top             =   810
                  Width           =   1140
               End
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  DataField       =   "wait_time"
                  Height          =   315
                  Index           =   24
                  Left            =   2250
                  MaxLength       =   50
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   11
                  Tag             =   "Partner"
                  Top             =   1815
                  Width           =   1140
               End
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  DataField       =   "cycle_time"
                  Height          =   315
                  Index           =   2
                  Left            =   2250
                  MaxLength       =   50
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   8
                  Tag             =   "Partner"
                  Top             =   300
                  Width           =   1140
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cycle Time (Sec)"
                  Height          =   195
                  Index           =   3
                  Left            =   615
                  TabIndex        =   64
                  Top             =   360
                  Width           =   1185
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Queue Time (Sec)"
                  Height          =   195
                  Index           =   4
                  Left            =   615
                  TabIndex        =   63
                  Top             =   1372
                  Width           =   1275
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Utilization ( % )"
                  Height          =   195
                  Index           =   5
                  Left            =   615
                  TabIndex        =   62
                  Top             =   2385
                  Width           =   1110
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Setup Time (Sec)"
                  Height          =   195
                  Index           =   26
                  Left            =   615
                  TabIndex        =   61
                  Top             =   866
                  Width           =   1215
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Wait Time (Sec)"
                  Height          =   195
                  Index           =   27
                  Left            =   615
                  TabIndex        =   60
                  Top             =   1878
                  Width           =   1125
               End
               Begin VB.Line Line1 
                  Index           =   5
                  X1              =   585
                  X2              =   2880
                  Y1              =   2625
                  Y2              =   2625
               End
               Begin VB.Line Line1 
                  Index           =   4
                  X1              =   585
                  X2              =   2880
                  Y1              =   1605
                  Y2              =   1605
               End
               Begin VB.Line Line1 
                  Index           =   3
                  X1              =   585
                  X2              =   2880
                  Y1              =   600
                  Y2              =   600
               End
               Begin VB.Line Line1 
                  Index           =   21
                  X1              =   585
                  X2              =   2880
                  Y1              =   1110
                  Y2              =   1110
               End
               Begin VB.Line Line1 
                  Index           =   22
                  X1              =   585
                  X2              =   2880
                  Y1              =   2115
                  Y2              =   2115
               End
            End
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "description"
               Height          =   330
               Index           =   1
               Left            =   1440
               MaxLength       =   50
               ScrollBars      =   2  'Vertical
               TabIndex        =   3
               Tag             =   "Partner"
               Top             =   435
               Width           =   3945
            End
            Begin VB.TextBox txtBox 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "Concurrent_Capacities"
               Height          =   330
               Index           =   6
               Left            =   7215
               MaxLength       =   10
               TabIndex        =   5
               Tag             =   "Partner"
               Top             =   75
               Width           =   1545
            End
            Begin VB.CommandButton cmdLink 
               Enabled         =   0   'False
               Height          =   315
               Index           =   3
               Left            =   9735
               Style           =   1  'Graphical
               TabIndex        =   58
               Top             =   510
               Visible         =   0   'False
               Width           =   345
            End
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "formID"
               Height          =   330
               Index           =   12
               Left            =   7215
               MaxLength       =   15
               TabIndex        =   57
               Tag             =   "Partner"
               Top             =   435
               Width           =   1545
            End
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "Capacity_Time_UOM_Code"
               Height          =   330
               Index           =   7
               Left            =   7215
               MaxLength       =   10
               TabIndex        =   56
               Top             =   795
               Visible         =   0   'False
               Width           =   1545
            End
            Begin VB.CommandButton cmdLink 
               Enabled         =   0   'False
               Height          =   315
               Index           =   4
               Left            =   8760
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   810
               Visible         =   0   'False
               Width           =   345
            End
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "WCID"
               Height          =   330
               Index           =   0
               Left            =   1440
               MaxLength       =   10
               TabIndex        =   2
               Tag             =   "Partner"
               Top             =   75
               Width           =   3945
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00EAAF6F&
               Caption         =   " Performance "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2235
               Left            =   4920
               TabIndex        =   50
               Top             =   1410
               Width           =   5445
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Height          =   315
                  Index           =   17
                  Left            =   1860
                  Locked          =   -1  'True
                  MaxLength       =   50
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   17
                  Top             =   1710
                  Width           =   1110
               End
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Height          =   315
                  Index           =   19
                  Left            =   1860
                  Locked          =   -1  'True
                  MaxLength       =   50
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   13
                  Top             =   690
                  Width           =   1110
               End
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Height          =   315
                  Index           =   18
                  Left            =   1860
                  Locked          =   -1  'True
                  MaxLength       =   50
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   15
                  Top             =   1200
                  Width           =   1110
               End
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Height          =   315
                  Index           =   20
                  Left            =   4140
                  Locked          =   -1  'True
                  MaxLength       =   50
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   18
                  Top             =   1710
                  Width           =   1110
               End
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Height          =   315
                  Index           =   21
                  Left            =   4140
                  Locked          =   -1  'True
                  MaxLength       =   50
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   16
                  Top             =   1200
                  Width           =   1110
               End
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Height          =   315
                  Index           =   22
                  Left            =   4140
                  Locked          =   -1  'True
                  MaxLength       =   50
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   14
                  Top             =   690
                  Width           =   1110
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Overhead"
                  Height          =   195
                  Index           =   6
                  Left            =   285
                  TabIndex        =   55
                  Top             =   1755
                  Width           =   720
               End
               Begin VB.Line Line1 
                  Index           =   17
                  X1              =   285
                  X2              =   4420
                  Y1              =   2010
                  Y2              =   2010
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Material"
                  Height          =   195
                  Index           =   7
                  Left            =   285
                  TabIndex        =   54
                  Top             =   1245
                  Width           =   570
               End
               Begin VB.Line Line1 
                  Index           =   18
                  X1              =   285
                  X2              =   4420
                  Y1              =   1500
                  Y2              =   1500
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tenaga Kerja"
                  Height          =   195
                  Index           =   8
                  Left            =   285
                  TabIndex        =   53
                  Top             =   735
                  Width           =   960
               End
               Begin VB.Line Line1 
                  Index           =   19
                  X1              =   285
                  X2              =   4420
                  Y1              =   990
                  Y2              =   990
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Per Hour"
                  Height          =   195
                  Index           =   24
                  Left            =   2085
                  TabIndex        =   52
                  Top             =   315
                  Width           =   630
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Per Unit"
                  Height          =   195
                  Index           =   25
                  Left            =   4395
                  TabIndex        =   51
                  Top             =   330
                  Width           =   570
               End
            End
            Begin MSDataListLib.DataCombo DataCombo1 
               DataField       =   "Calendar"
               Height          =   315
               Left            =   1440
               TabIndex        =   4
               Tag             =   "Partner"
               Top             =   795
               Width           =   3945
               _ExtentX        =   6959
               _ExtentY        =   714
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               ListField       =   "Description"
               BoundColumn     =   "Calendar"
               Text            =   "DataCombo1"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Keterangan"
               Height          =   195
               Index           =   0
               Left            =   195
               TabIndex        =   70
               Top             =   510
               Width           =   840
            End
            Begin VB.Line Line1 
               Index           =   2
               X1              =   195
               X2              =   1620
               Y1              =   750
               Y2              =   750
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Calendar ID"
               Height          =   195
               Index           =   2
               Left            =   195
               TabIndex        =   69
               Top             =   863
               Width           =   855
            End
            Begin VB.Line Line1 
               Index           =   1
               X1              =   195
               X2              =   1620
               Y1              =   1095
               Y2              =   1095
            End
            Begin VB.Line Line1 
               Index           =   6
               X1              =   5565
               X2              =   7230
               Y1              =   390
               Y2              =   390
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Concurrent Capacities"
               Height          =   195
               Index           =   9
               Left            =   5580
               TabIndex        =   68
               Top             =   150
               Width           =   1590
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Form ID"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   22
               Left            =   5580
               TabIndex        =   67
               Top             =   510
               Width           =   570
            End
            Begin VB.Line Line1 
               Index           =   20
               X1              =   5565
               X2              =   7230
               Y1              =   750
               Y2              =   750
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Work Center"
               Height          =   195
               Index           =   1
               Left            =   195
               TabIndex        =   66
               Top             =   143
               Width           =   915
            End
            Begin VB.Line Line1 
               Index           =   7
               Visible         =   0   'False
               X1              =   5565
               X2              =   7230
               Y1              =   1110
               Y2              =   1110
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Capacity TIME"
               Height          =   195
               Index           =   10
               Left            =   5580
               TabIndex        =   65
               Top             =   870
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Line Line1 
               Index           =   0
               X1              =   195
               X2              =   1620
               Y1              =   390
               Y2              =   390
            End
         End
         Begin VB.PictureBox Picture5 
            Height          =   4575
            Left            =   100
            ScaleHeight     =   4515
            ScaleWidth      =   10500
            TabIndex        =   48
            Top             =   400
            Width           =   10560
            Begin MSComctlLib.ListView ListView1 
               Height          =   4500
               Left            =   0
               TabIndex        =   1
               Top             =   -15
               Width           =   10515
               _ExtentX        =   18547
               _ExtentY        =   7938
               View            =   3
               Sorted          =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               AllowReorder    =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               HotTracking     =   -1  'True
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
               NumItems        =   8
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "WCID"
                  Object.Width           =   2822
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Keterangan"
                  Object.Width           =   4762
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Calendar"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   3
                  Text            =   "Cycle Time"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   4
                  Text            =   "Queue Time"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Wait Time"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "Setup Time"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   7
                  Text            =   "Utilization"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.PictureBox Picture4 
            Height          =   4545
            Index           =   1
            Left            =   -74900
            ScaleHeight     =   4485
            ScaleWidth      =   10500
            TabIndex        =   46
            Top             =   400
            Width           =   10560
            Begin MSDataGridLib.DataGrid DataGrid1 
               Height          =   4515
               Index           =   0
               Left            =   -15
               TabIndex        =   47
               Top             =   -15
               Width           =   10530
               _ExtentX        =   18574
               _ExtentY        =   7964
               _Version        =   393216
               AllowUpdate     =   0   'False
               Appearance      =   0
               HeadLines       =   1
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
               ColumnCount     =   4
               BeginProperty Column00 
                  DataField       =   "IDX"
                  Caption         =   "IDX"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   "WCID"
                  Caption         =   "WCID"
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
                  DataField       =   "StageID"
                  Caption         =   "StageID"
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
                  DataField       =   "no"
                  Caption         =   "No"
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
         Begin VB.PictureBox Picture4 
            Height          =   4545
            Index           =   0
            Left            =   -74900
            ScaleHeight     =   4485
            ScaleWidth      =   10500
            TabIndex        =   44
            Top             =   400
            Width           =   10560
            Begin MSDataListLib.DataList ListResource 
               Height          =   1950
               Left            =   4905
               TabIndex        =   74
               Top             =   1290
               Visible         =   0   'False
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   3440
               _Version        =   393216
               Appearance      =   0
            End
            Begin VB.ListBox ListUOM 
               Appearance      =   0  'Flat
               Height          =   1080
               ItemData        =   "FrmWCTrans.frx":6CBC
               Left            =   0
               List            =   "FrmWCTrans.frx":6CC9
               TabIndex        =   73
               Top             =   0
               Visible         =   0   'False
               Width           =   1260
            End
            Begin VB.ListBox ListSatuanWaktu 
               Appearance      =   0  'Flat
               Height          =   1080
               ItemData        =   "FrmWCTrans.frx":6CE3
               Left            =   0
               List            =   "FrmWCTrans.frx":6CF0
               TabIndex        =   72
               Top             =   0
               Visible         =   0   'False
               Width           =   1260
            End
            Begin VB.ListBox ListTypeID 
               Appearance      =   0  'Flat
               Height          =   1080
               ItemData        =   "FrmWCTrans.frx":6D0A
               Left            =   2115
               List            =   "FrmWCTrans.frx":6D17
               TabIndex        =   71
               Top             =   1230
               Visible         =   0   'False
               Width           =   1260
            End
            Begin MSDataGridLib.DataGrid DataGrid1 
               Height          =   4515
               Index           =   1
               Left            =   -15
               TabIndex        =   45
               Top             =   -30
               Width           =   10545
               _ExtentX        =   18600
               _ExtentY        =   7964
               _Version        =   393216
               AllowUpdate     =   0   'False
               Appearance      =   0
               HeadLines       =   1
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
                  DataField       =   "ID"
                  Caption         =   "ID"
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
                  DataField       =   "WCID"
                  Caption         =   "WCID"
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
                  DataField       =   "TypeID"
                  Caption         =   "TypeID"
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
                  DataField       =   "capacity_qty"
                  Caption         =   "JML Kapasitas"
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
                  DataField       =   "capacity_uom"
                  Caption         =   "Satuan Kapasitas"
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
               BeginProperty Column05 
                  DataField       =   "capacity_time_uom"
                  Caption         =   "Kapasitas per Satuan Waktu"
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
                  DataField       =   "resourceName"
                  Caption         =   "Resource Type"
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
                     Button          =   -1  'True
                     Locked          =   -1  'True
                  EndProperty
                  BeginProperty Column03 
                  EndProperty
                  BeginProperty Column04 
                     Button          =   -1  'True
                     Locked          =   -1  'True
                  EndProperty
                  BeginProperty Column05 
                     Button          =   -1  'True
                     Locked          =   -1  'True
                  EndProperty
                  BeginProperty Column06 
                     Button          =   -1  'True
                     Locked          =   -1  'True
                  EndProperty
               EndProperty
            End
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00EAAF6F&
            Height          =   4545
            Left            =   -74895
            ScaleHeight     =   4485
            ScaleWidth      =   10500
            TabIndex        =   31
            Top             =   405
            Width           =   10560
            Begin VB.Frame Frame3 
               BackColor       =   &H00EAAF6F&
               Caption         =   " Jadwal Kerja "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1620
               Left            =   90
               TabIndex        =   33
               Top             =   2520
               Width           =   10290
               Begin VB.CheckBox ChkDays 
                  BackColor       =   &H00EAAF6F&
                  DataSource      =   "MyDDE"
                  Height          =   225
                  Index           =   6
                  Left            =   8310
                  TabIndex        =   26
                  Tag             =   "partner"
                  Top             =   915
                  Width           =   255
               End
               Begin VB.CheckBox ChkDays 
                  BackColor       =   &H00EAAF6F&
                  DataSource      =   "MyDDE"
                  Height          =   225
                  Index           =   5
                  Left            =   7198
                  TabIndex        =   25
                  Tag             =   "partner"
                  Top             =   915
                  Width           =   255
               End
               Begin VB.CheckBox ChkDays 
                  BackColor       =   &H00EAAF6F&
                  DataSource      =   "MyDDE"
                  Height          =   225
                  Index           =   4
                  Left            =   6087
                  TabIndex        =   24
                  Tag             =   "partner"
                  Top             =   915
                  Width           =   255
               End
               Begin VB.CheckBox ChkDays 
                  BackColor       =   &H00EAAF6F&
                  DataSource      =   "MyDDE"
                  Height          =   225
                  Index           =   3
                  Left            =   4976
                  TabIndex        =   23
                  Tag             =   "partner"
                  Top             =   915
                  Width           =   255
               End
               Begin VB.CheckBox ChkDays 
                  BackColor       =   &H00EAAF6F&
                  DataSource      =   "MyDDE"
                  Height          =   225
                  Index           =   2
                  Left            =   3865
                  TabIndex        =   22
                  Tag             =   "partner"
                  Top             =   915
                  Width           =   255
               End
               Begin VB.CheckBox ChkDays 
                  BackColor       =   &H00EAAF6F&
                  DataSource      =   "MyDDE"
                  Height          =   225
                  Index           =   1
                  Left            =   2754
                  TabIndex        =   21
                  Tag             =   "partner"
                  Top             =   915
                  Width           =   255
               End
               Begin VB.CheckBox ChkDays 
                  BackColor       =   &H00EAAF6F&
                  DataField       =   "mon_days"
                  DataSource      =   "MyDDE"
                  Height          =   225
                  Index           =   0
                  Left            =   1643
                  TabIndex        =   20
                  Tag             =   "partner"
                  Top             =   915
                  Width           =   255
               End
               Begin VB.Label LblDays 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Sabtu"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   7
                  Left            =   6750
                  TabIndex        =   42
                  Top             =   435
                  Width           =   1125
               End
               Begin VB.Label LblDays 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Jumat"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   8
                  Left            =   5625
                  TabIndex        =   41
                  Top             =   435
                  Width           =   1125
               End
               Begin VB.Label LblDays 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Kamis"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   9
                  Left            =   4500
                  TabIndex        =   40
                  Top             =   435
                  Width           =   1125
               End
               Begin VB.Label LblDays 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Rabu"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   10
                  Left            =   3375
                  TabIndex        =   39
                  Top             =   435
                  Width           =   1125
               End
               Begin VB.Label LblDays 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Selasa"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   11
                  Left            =   2250
                  TabIndex        =   38
                  Top             =   435
                  Width           =   1125
               End
               Begin VB.Label LblDays 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Senin"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   12
                  Left            =   1125
                  TabIndex        =   37
                  Top             =   435
                  Width           =   1125
               End
               Begin VB.Label LblDays 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Minggu"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   13
                  Left            =   7875
                  TabIndex        =   36
                  Top             =   435
                  Width           =   1125
               End
               Begin VB.Label LblJadwal 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EAAF6F&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Hari"
                  Height          =   210
                  Index           =   0
                  Left            =   330
                  TabIndex        =   35
                  Top             =   457
                  Width           =   300
               End
               Begin VB.Label LblJadwal 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EAAF6F&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Status"
                  Height          =   210
                  Index           =   1
                  Left            =   330
                  TabIndex        =   34
                  Top             =   885
                  Width           =   525
               End
            End
            Begin MSComCtl2.DTPicker TimeStart 
               DataSource      =   "MyDDE"
               Height          =   315
               Index           =   7
               Left            =   2745
               TabIndex        =   32
               Tag             =   "Partner"
               Top             =   900
               Visible         =   0   'False
               Width           =   1290
               _ExtentX        =   2275
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
               Format          =   65863682
               UpDown          =   -1  'True
               CurrentDate     =   38467
            End
            Begin MSDataGridLib.DataGrid DataGrid1 
               Height          =   2385
               Index           =   2
               Left            =   -15
               TabIndex        =   43
               Top             =   -15
               Width           =   10515
               _ExtentX        =   18547
               _ExtentY        =   4207
               _Version        =   393216
               AllowUpdate     =   0   'False
               Appearance      =   0
               HeadLines       =   1
               RowHeight       =   19
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
               ColumnCount     =   14
               BeginProperty Column00 
                  DataField       =   "ID"
                  Caption         =   "ID"
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
                  DataField       =   "WCID"
                  Caption         =   "WCID"
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
                  DataField       =   "ShiftID"
                  Caption         =   "ShiftID"
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
                  DataField       =   "Shift_desc"
                  Caption         =   "Shift_desc"
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
                  DataField       =   "start_time"
                  Caption         =   "start_time"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "h:mm:ss AMPM"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   4
                  EndProperty
               EndProperty
               BeginProperty Column05 
                  DataField       =   "stop_time"
                  Caption         =   "stop_time"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "h:mm:ss AMPM"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   4
                  EndProperty
               EndProperty
               BeginProperty Column06 
                  DataField       =   "break_time"
                  Caption         =   "break_time"
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
               BeginProperty Column07 
                  DataField       =   "mon_days"
                  Caption         =   "mon_days"
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
               BeginProperty Column08 
                  DataField       =   "tue_days"
                  Caption         =   "tue_days"
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
               BeginProperty Column09 
                  DataField       =   "wed_days"
                  Caption         =   "wed_days"
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
               BeginProperty Column10 
                  DataField       =   "thu_days"
                  Caption         =   "thu_days"
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
               BeginProperty Column11 
                  DataField       =   "fri_days"
                  Caption         =   "fri_days"
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
               BeginProperty Column12 
                  DataField       =   "sat_days"
                  Caption         =   "sat_days"
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
               BeginProperty Column13 
                  DataField       =   "sun_days"
                  Caption         =   "sun_days"
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
                     Locked          =   -1  'True
                  EndProperty
                  BeginProperty Column03 
                  EndProperty
                  BeginProperty Column04 
                  EndProperty
                  BeginProperty Column05 
                  EndProperty
                  BeginProperty Column06 
                  EndProperty
                  BeginProperty Column07 
                  EndProperty
                  BeginProperty Column08 
                  EndProperty
                  BeginProperty Column09 
                  EndProperty
                  BeginProperty Column10 
                  EndProperty
                  BeginProperty Column11 
                  EndProperty
                  BeginProperty Column12 
                  EndProperty
                  BeginProperty Column13 
                  EndProperty
               EndProperty
            End
         End
         Begin MSDataGridLib.DataGrid GrdBOMWhere 
            Height          =   4660
            Left            =   -74985
            TabIndex        =   75
            Tag             =   "SL"
            Top             =   345
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   8229
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16577005
            ForeColor       =   7159830
            HeadLines       =   2
            RowHeight       =   16
            RowDividerStyle =   6
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
            ColumnCount     =   6
            BeginProperty Column00 
               DataField       =   "WCID"
               Caption         =   "BOM Reference"
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
               DataField       =   "BomReff"
               Caption         =   "Description"
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
               DataField       =   "ItemName"
               Caption         =   "Item Barang"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "dd MMMM yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "ResourcesID"
               Caption         =   "Resources ID"
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
               DataField       =   "StageNote"
               Caption         =   "Stage Note"
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
               DataField       =   "NoLine"
               Caption         =   "NoLine"
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
               MarqueeStyle    =   4
               BeginProperty Column00 
                  DividerStyle    =   6
               EndProperty
               BeginProperty Column01 
                  DividerStyle    =   6
               EndProperty
               BeginProperty Column02 
               EndProperty
               BeginProperty Column03 
               EndProperty
               BeginProperty Column04 
               EndProperty
               BeginProperty Column05 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid GrdMOUsed 
            Height          =   4620
            Left            =   -74990
            TabIndex        =   76
            Tag             =   "SL"
            Top             =   360
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   8149
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   16577005
            ForeColor       =   7159830
            HeadLines       =   2
            RowHeight       =   16
            RowDividerStyle =   6
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
            ColumnCount     =   7
            BeginProperty Column00 
               DataField       =   "ItemName"
               Caption         =   "Item Barang"
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
               DataField       =   "SeqNo"
               Caption         =   "No Sequensial"
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
               DataField       =   "StartDate"
               Caption         =   "Mulai"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "dd MMMM yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "EndDate"
               Caption         =   "Selesai"
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
               DataField       =   "StatusMO"
               Caption         =   "Status MO"
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
               DataField       =   "CompanyName"
               Caption         =   "CompanyName"
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
               DataField       =   "Description"
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
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   4
               BeginProperty Column00 
                  DividerStyle    =   6
               EndProperty
               BeginProperty Column01 
                  DividerStyle    =   6
               EndProperty
               BeginProperty Column02 
               EndProperty
               BeginProperty Column03 
               EndProperty
               BeginProperty Column04 
               EndProperty
               BeginProperty Column05 
               EndProperty
               BeginProperty Column06 
               EndProperty
            EndProperty
         End
         Begin CrystalActiveXReportViewerLib11Ctl.CrystalActiveXReportViewer CrWCInQuiry 
            Height          =   4575
            Left            =   -74985
            TabIndex        =   78
            Top             =   360
            Width           =   9345
            _cx             =   16484
            _cy             =   8070
            DisplayGroupTree=   0   'False
            DisplayToolbar  =   0   'False
            EnableGroupTree =   -1  'True
            EnableNavigationControls=   -1  'True
            EnableStopButton=   -1  'True
            EnablePrintButton=   -1  'True
            EnableZoomControl=   -1  'True
            EnableCloseButton=   0   'False
            EnableProgressControl=   0   'False
            EnableSearchControl=   -1  'True
            EnableRefreshButton=   -1  'True
            EnableDrillDown =   -1  'True
            EnableAnimationControl=   0   'False
            EnableSelectExpertButton=   0   'False
            EnableToolbar   =   -1  'True
            DisplayBorder   =   0   'False
            DisplayTabs     =   0   'False
            DisplayBackgroundEdge=   -1  'True
            SelectionFormula=   ""
            EnablePopupMenu =   0   'False
            EnableExportButton=   -1  'True
            EnableSearchExpertButton=   0   'False
            EnableHelpButton=   0   'False
            LaunchHTTPHyperlinksInNewBrowser=   -1  'True
            EnableLogonPrompts=   -1  'True
            LocaleID        =   1057
         End
         Begin VB.Label lblChartType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chart Type"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   -65520
            TabIndex        =   77
            Top             =   360
            Width           =   915
         End
      End
   End
End
Attribute VB_Name = "FrmWorkCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RcCal As New DBQuick
Private RcDetail As New DBQuick
Private RcStage As New DBQuick
Private WithEvents RcShift As DBQuick
Attribute RcShift.VB_VarHelpID = -1
Private MEdit, mFirstCaller As Boolean
Private RcPartner As New DBQuick
Private colActive As Integer
Private RsResourceType As New DBQuick
Private RcBOMWhere As New DBQuick
Private RcMOWhere As New DBQuick
Private Rc As New DBQuick
Public m_Report As CRAXDRT.Report
Private m_Application As New CRAXDRT.Application

Private Sub ChkDays_Click(Index As Integer)
    Select Case Index
        Case 0: RcShift.DBRecordset.Fields("mon_days") = ChkDays(Index).Value
        Case 1: RcShift.DBRecordset.Fields("tue_days") = ChkDays(Index).Value
        Case 2: RcShift.DBRecordset.Fields("wed_days") = ChkDays(Index).Value
        Case 3: RcShift.DBRecordset.Fields("thu_days") = ChkDays(Index).Value
        Case 4: RcShift.DBRecordset.Fields("fri_days") = ChkDays(Index).Value
        Case 5: RcShift.DBRecordset.Fields("sat_days") = ChkDays(Index).Value
        Case 6: RcShift.DBRecordset.Fields("sun_days") = ChkDays(Index).Value
    End Select
End Sub

Private Sub LoadUOM()
    Dim RsUOM As New DBQuick
    RsUOM.DBOpen "SELECT * FROM [UOM Table] ORDER BY UOM", CNN
    ListUOM.Clear
    If RsUOM.DBRecordset.Recordcount > 0 Then
        While Not RsUOM.DBRecordset.EOF
            ListUOM.AddItem RsUOM.DBRecordset.Fields(0)
            RsUOM.DBRecordset.MoveNext
        Wend
    End If
    RsUOM.CloseDB
End Sub
Private Sub CountPerformance()
    txtBox(17).Text = ""
    txtBox(18).Text = ""
    txtBox(19).Text = ""
    txtBox(20).Text = ""
    txtBox(21).Text = ""
    txtBox(22).Text = ""
    
    With RcDetail.DBRecordset
    If .Recordcount > 0 Then
        .MoveFirst
        While Not .EOF
            Select Case UCase(.Fields("TypeID"))
                Case "LABOR":
                    txtBox(22).Text = .Fields("Capacity_qty")
                    Select Case UCase(.Fields("capacity_time_uom"))
                        Case "DAYS": txtBox(19).Text = Val(.Fields("Capacity_qty")) / 24
                        Case "HOURS": txtBox(19).Text = .Fields("Capacity_qty")
                        Case "MINUTES": txtBox(19).Text = Val(.Fields("Capacity_qty")) * 60
                    End Select
                Case "MATERIAL":
                    txtBox(21).Text = .Fields("Capacity_qty")
                    Select Case UCase(.Fields("capacity_time_uom"))
                        Case "DAYS": txtBox(18).Text = Val(.Fields("Capacity_qty")) / 24
                        Case "HOURS": txtBox(18).Text = .Fields("Capacity_qty")
                        Case "MINUTES": txtBox(18).Text = Val(.Fields("Capacity_qty")) * 60
                    End Select
                Case "OVERHEAD":
                    txtBox(20).Text = .Fields("Capacity_qty")
                    Select Case UCase(.Fields("capacity_time_uom"))
                        Case "DAYS": txtBox(17).Text = Val(.Fields("Capacity_qty")) / 24
                        Case "HOURS": txtBox(17).Text = .Fields("Capacity_qty")
                        Case "MINUTES": txtBox(17).Text = Val(.Fields("Capacity_qty")) * 60
                    End Select
            End Select
            .MoveNext
        Wend
    End If
    End With
End Sub


Private Sub cmdLink_Click(Index As Integer)
    OpenPartner Index
End Sub

Private Sub DataCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub

Private Sub DataGrid1_ButtonClick(Index As Integer, ByVal ColIndex As Integer)
   If MEdit Then
       Select Case Index
           Case 1:
               Select Case ColIndex
                   Case 6:
                      LoadResourceTable
                      ListResource.Move DataGrid1(Index).Columns(ColIndex).Left, _
                                             DataGrid1(Index).RowTop(DataGrid1(Index).row) + 200, _
                                             DataGrid1(Index).Columns(ColIndex).width
                      ListResource.Visible = True
                   Case 5:
                       ListSatuanWaktu.Move DataGrid1(Index).Columns(ColIndex).Left, _
                                             DataGrid1(Index).RowTop(DataGrid1(Index).row) + 200, _
                                             DataGrid1(Index).Columns(ColIndex).width
                       ListSatuanWaktu.Visible = True
                   Case 4:
                       ListUOM.Move DataGrid1(Index).Columns(ColIndex).Left, _
                                             DataGrid1(Index).RowTop(DataGrid1(Index).row) + 200, _
                                             DataGrid1(Index).Columns(ColIndex).width
                       ListUOM.Visible = True
                   Case 2:
                       ListTypeID.Move DataGrid1(Index).Columns(ColIndex).Left, _
                                             DataGrid1(Index).RowTop(DataGrid1(Index).row) + 200, _
                                             DataGrid1(Index).Columns(ColIndex).width
                       ListTypeID.Visible = True
               End Select
      End Select
   End If
End Sub

Private Sub DataGrid1_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
   ListUOM.Visible = False
   ListSatuanWaktu.Visible = False
   ListResource.Visible = False
   ListTypeID.Visible = False
    Select Case Index
        Case 2:
            If ((DataGrid1(Index).col = 4) Or (DataGrid1(Index).col = 5)) And MEdit Then
                If DataGrid1(Index).col = 4 Then
                    If Not IsNull(RcShift.DBRecordset.Fields("start_time")) Then TimeStart(7).Value = RcShift.DBRecordset.Fields("start_time")
                Else
                    If Not IsNull(RcShift.DBRecordset.Fields("stop_time")) Then TimeStart(7).Value = RcShift.DBRecordset.Fields("stop_time")
                End If
                TimeStart(7).Visible = True
                TimeStart(7).Move DataGrid1(Index).Columns(DataGrid1(Index).col).Left, _
                                  DataGrid1(Index).RowTop(DataGrid1(Index).row), _
                                  DataGrid1(Index).Columns(DataGrid1(Index).col).width, _
                                  DataGrid1(Index).RowHeight
                TimeStart(7).SetFocus
                
            End If
    End Select
    colActive = DataGrid1(Index).col
End Sub

Private Sub Form_Load()
    Set RcShift = New DBQuick
    HiasFormManTell Picture2, Me
    SSTab1.Tab = 0
    OpenCalendar
    With MyDDE
        .EditModeReplace = False
        Set .BindForm = Me
        .BindFormTAG = "Partner"
        Set .ActiveConnection = CNN
        .PrepareQuery = "SELECT WCID, Description AS Keterangan, CalendarID AS Calendar, " & _
                " cycle_time, queue_time, setup_time, wait_time, utilization, " & _
                " Concurrent_Capacities, Capacity_UOM, Capacity_Time_UOM_Code,Description,formID FROM wcenter_header ORDER BY WCID"
    End With
    OpenHeader
    OpenChild
    CountPerformance
    LoadUOM
    Set mCall = New frmCaller
    DataGrid1(2).RowHeight = 250
End Sub

Private Sub OpenChild()
    OpenDetail txtBox(0)
    OpenStage txtBox(0)
    Openshift txtBox(0)
End Sub

Private Sub Openshift(Params As String)
    RcShift.DBOpen "SELECT * FROM wcenter_shift  WHERE (wcenter_shift.WCID = N'" & Params & "') ORDER BY shiftID", CNN, lckLockBatch
    Set DataGrid1(2).DataSource = RcShift.DBRecordset
End Sub

Private Sub PrepareSQL()
With MyDDE
    With MyDDE
        .PrepareAppend = "insert into wcenter_header ([WCID] ,[CalendarID],[Description],[cycle_time],[queue_time]," & _
                                 " [setup_time],[wait_time],[utilization],[Maximum_Efficiency]," & _
                                 " [Minimum_Efficiency],[Blocked],[Fixed_Scrap_Quantity],[Scrap]," & _
                                 " [cycle_time_uom_code],[Queue_Time_UOM_Code],[Setup_Time_UOM_Code],[Wait_Time_UOM_Code]," & _
                                 " [Flushing_Method],[No_Series],[Overhead_Rate],[Gen_Prod_Posting_Group],formID,Concurrent_Capacities) " & _
                         "values ('" & .GetFieldByName("WCID") & "','" & _
                                 DataCombo1.BoundText & "','" & _
                                 .GetFieldByName("Description") & "'," & _
                                 FQty(.GetFieldByName("cycle_time")) & "," & _
                                 FQty(.GetFieldByName("queue_time")) & "," & _
                                 FQty(.GetFieldByName("setup_time")) & "," & _
                                 FQty(.GetFieldByName("wait_time")) & "," & _
                                 FQty(.GetFieldByName("utilization")) & "," & _
                                 FQty(.GetFieldByName("Maximum_Efficiency")) & "," & _
                                 FQty(.GetFieldByName("Maximum_Efficiency")) & "," & _
                                 FQty(.GetFieldByName("Blocked")) & "," & _
                                 FQty(.GetFieldByName("Fixed_Scrap_Quantity")) & "," & _
                                 FQty(.GetFieldByName("scrap")) & ",'" & _
                                 .GetFieldByName("cycle_time_uom_code") & "','" & _
                                 .GetFieldByName("queue_time_uom_code") & "','" & _
                                 .GetFieldByName("setup_time_uom_code") & "','" & _
                                 .GetFieldByName("wait_time_UOM_code") & "'," & _
                                 FQty(.GetFieldByName("flushing_method")) & ",'" & _
                                 .GetFieldByName("no_series") & "'," & FQty(.GetFieldByName("Overhead_rate")) & _
                                 ",'" & .GetFieldByName("gen_prod_posting_group") & "'," & FQty(txtBox(12).Text) & "," & FQty(txtBox(6).Text) & ")"
   
        .PrepareUpdate = " update wcenter_header set [CalendarID] ='" & DataCombo1.BoundText & "'," & _
                              "[Description]='" & .GetFieldByName("Description") & "'," & _
                              "[cycle_time] = " & FQty(.GetFieldByName("cycle_time")) & "," & _
                              "[queue_time] = " & FQty(.GetFieldByName("queue_time")) & "," & _
                              "[setup_time] = " & FQty(.GetFieldByName("setup_time")) & "," & _
                              "[wait_time]  = " & FQty(.GetFieldByName("wait_time")) & "," & _
                              "[utilization]= " & FQty(.GetFieldByName("utilization")) & "," & _
                              "[Maximum_Efficiency] = " & FQty(.GetFieldByName("Maximum_Efficiency")) & "," & _
                              "[Minimum_Efficiency] = " & FQty(.GetFieldByName("Minimum_Efficiency")) & "," & _
                              "[Blocked] = " & FQty(.GetFieldByName("Blocked")) & "," & _
                              "[Fixed_Scrap_Quantity] = " & FQty(.GetFieldByName("Fixed_Scrap_Quantity")) & "," & _
                              "[Scrap] = " & FQty(.GetFieldByName("scrap")) & "," & _
                              "[cycle_time_uom_code] = '" & .GetFieldByName("cycle_time_uom_code") & "'," & _
                              "[Queue_Time_UOM_Code] = '" & .GetFieldByName("queue_time_uom_code") & "'," & _
                              "[Setup_Time_UOM_Code] = '" & .GetFieldByName("setup_time_uom_code") & "'," & _
                              "[Wait_Time_UOM_Code] = '" & .GetFieldByName("wait_time_UOM_code") & "'," & _
                              "[Flushing_Method] = " & FQty(.GetFieldByName("flushing_method")) & "," & _
                              "[No_Series] = '" & .GetFieldByName("no_series") & "'," & _
                              "[Overhead_Rate] = " & FQty(.GetFieldByName("Overhead_rate")) & "," & _
                              "[Gen_Prod_Posting_Group] ='" & .GetFieldByName("gen_prod_posting_group") & "'," & _
                              "formID =" & FQty(txtBox(12).Text) & "," & _
                              "Concurrent_Capacities =" & FQty(txtBox(6).Text) & _
                       " where WCID ='" & .GetFieldByName("WCID") & "'"
                     
        .PrepareDelete = " delete from wcenter_header where WCID ='" & .GetFieldByName("WCID") & "'"
                     
    End With
End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ScanKey KeyCode, Shift, MyDDE
End Sub


Private Sub LoadResourceTable()
   RsResourceType.DBOpen "select ResourcesID,description from [Resources Table] where typeID='" & MyDDE.ChildRecordset.Fields("TypeID") & "'", CNN
   ListResource.DataField = "ResourcesID"
   ListResource.ListField = "description"
   Set ListResource.RowSource = RsResourceType.DBRecordset
End Sub

Private Sub OpenHeader()
Dim Rc As New DBQuick
Dim I As Integer
Dim Avdata As Variant
Set Rc.DBRecordset = MyDDE.ActiveRecordset.Clone(adLockReadOnly)
ListView1.ListItems.Clear
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


Private Sub ListResource_Click()
   MyDDE.ChildRecordset.Fields("ResourceName") = ListResource.Text
   If RsResourceType.DBRecordset.Recordcount > 0 Then
      MyDDE.ChildRecordset.Fields("Resource_type") = RsResourceType.DBRecordset.Fields("ResourcesID")
   End If
   ListResource.Visible = False
End Sub

Private Sub ListSatuanWaktu_Click()
   ListSatuanWaktu.Visible = False
   RcDetail.DBRecordset.Fields("Capacity_time_uom") = ListSatuanWaktu.Text
   DataGrid1(1).Columns(5).Text = ListSatuanWaktu.Text
End Sub

Private Sub ListTypeID_Click()
   ListTypeID.Visible = False
   RcDetail.DBRecordset.Fields("TypeID") = ListTypeID.Text
   DataGrid1(1).Columns(2).Text = ListTypeID.Text
End Sub

Private Sub ListUOM_Click()
    ListUOM.Visible = False
    RcDetail.DBRecordset.Fields("Capacity_uom") = ListUOM.Text
    DataGrid1(1).Columns(4).Text = ListUOM.Text
End Sub

Private Sub ListView1_DblClick()
    SSTab1.Tab = 1
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
If MyDDE.ActiveRecordset.Recordcount <> 0 Then
   MyDDE.FindStringData "[WCID]='" & Item.Text & "'"
End If
End Sub


Private Sub mCall_BeforeUnload()
   Dim StrFlt As String
   Select Case UCase(mCall.FromTagActive)
       Case "DATA STAGE": StrFlt = "StageID"
       Case "CAPACITY UNIT"
           If txtBox(12).Enabled = True Then txtBox(12).SetFocus
           Exit Sub
       Case "CAPACITY TIME UNIT"
           If txtBox(7).Enabled = True Then txtBox(7).SetFocus
           Exit Sub
   End Select
   If UCase(mCall.FromTagActive) <> "DATA FORM" Then
      'Debug.Print MyDDE.ChildRecordset.Source
      If FindOwnRecordset(MyDDE.ChildRecordset, StrFlt & "= '" & MyDDE.ChildRecordset.Fields(StrFlt) & "'") = True Then
         MessageBox "Record -> " & MyDDE.ChildRecordset.Fields(StrFlt) & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
         MyDDE.ChildRecordset.CancelBatch adAffectCurrent
         If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
      Else
         If Not IsNull(MyDDE.ChildRecordset.Fields(StrFlt)) = True Then
            If MyDDE.ChildRecordset.Fields(StrFlt) = "" Then
               MyDDE.ChildRecordset.CancelBatch adAffectCurrent
               If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
               mFirstCaller = False
            End If
         End If
      End If
   End If
   Select Case SSTab1.Tab
          Case 3: DataGrid1(0).SetFocus
   End Select
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case UCase(TagForm)
       Case "DATA RESOURCES":
            With RcDetail.DBRecordset
                 .Fields("typeID") = mCall.GetFieldByName("resources")
                 mFirstCaller = False
            End With

       Case "DATA STAGE":
            With RcStage.DBRecordset
                 .Fields("WCID") = txtBox(0)
                 .Fields("StageID") = mCall.GetFieldByName("StageID")
                 
                 mFirstCaller = False
            End With

       Case "CAPACITY UNIT":
            With MyDDE
                 '.GetFieldByName("Capacity_UOM") = mCall.GetFieldByName(0)
                 mFirstCaller = False
            End With
            
       Case "CAPACITY TIME UNIT":
            With MyDDE
'                 .GetFieldByName("Capacity_Time_UOM_Code") = mCall.GetFieldByName(0)
                 mFirstCaller = False
            End With
       
       Case "DATA FORM":
            MyDDE.GetFieldByName("formID") = mCall.GetFieldByName(0)
            txtBox(12).Text = mCall.GetFieldByName(0)
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo AfterErr
MEdit = False
DataGrid1(0).AllowUpdate = False
DataGrid1(1).AllowUpdate = False
DataGrid1(2).AllowUpdate = False

Select Case AdReasonActiveDb
       Case tmbEdit:
            MEdit = True
            If SSTab1.Tab = 0 Then SSTab1.Tab = 1
            txtBox(0).Enabled = False
            txtBox(1).SetFocus
            DataGrid1(0).AllowUpdate = True
            DataGrid1(1).AllowUpdate = True
            DataGrid1(2).AllowUpdate = True

       Case tmbAddNew:
            MEdit = True
            MyDDE.GetFieldByName("WCID") = "New Order"
            MyDDE.GetFieldByName("Keterangan") = "-"
            SSTab1.Tab = 1
            txtBox(0).Enabled = True
            txtBox(1).SetFocus
            DataGrid1(0).AllowUpdate = True
            DataGrid1(1).AllowUpdate = True
            DataGrid1(2).AllowUpdate = True
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
                SavingResource
                SavingShift
                SavingStage
               MEdit = False
            End If
       Case tmbDelete:
       Case tmbDetail:
            MEdit = True
      '      If mFirstCaller = False Then
               DataGrid1(0).AllowUpdate = True
               DataGrid1(1).AllowUpdate = True
               DataGrid1(2).AllowUpdate = True
               MEdit = True
               Select Case SSTab1.Tab
                      Case 2:
                           Dim rloop As Boolean
                           Dim rCounter As Integer
                           If RcShift.Recordcount > 0 Then
                              RcShift.DBRecordset.MoveFirst
                              rCounter = 1
                              rloop = True
                              While rloop
                                   If RcShift.DBRecordset.Recordcount = RcShift.DBRecordset.Bookmark Then rloop = False
                                   RcShift.DBRecordset.Fields("shiftID") = rCounter
                                   RcShift.DBRecordset.MoveNext
                                   rCounter = rCounter + 1
                              Wend
                           End If
                      Case 3: OpenPartner 2
                      'Case 4: OpenPartner 1
               End Select
          '  End If
       Case tmbPrint:
            CallRPTReport "Manufacture WC Table.rpt", "Select * From [Manufacture WC Table] where [Work ID] ='" & txtBox(0) & "'"
       Case tmbQuit:
            Unload Me
            Set MyDDE.BindForm = Nothing
End Select
cmdLink(3).Enabled = MEdit
cmdLink(4).Enabled = MEdit
Exit Sub
AfterErr:
   MessageBox Err.Description, FrmWorkCenter.Caption, msgOkOnly, msgExclamation
   Err.Clear
End Sub

Private Sub SavingResource()
    With RcDetail.DBRecordset
         If .Recordcount <> 0 Then
            .MoveFirst
            If SendDataToServer("Delete From [wcenter_resources] where WCID=N'" & txtBox(0) & "'") = True Then
            Do
              If .EOF Then Exit Do
                 SendDataToServer (" INSERT INTO [wcenter_resources]" & _
                                   " (ID,WCID, TypeID, capacity_qty,capacity_uom,capacity_time_uom,resource_type)" & _
                                   " VALUES (newID(),N'" & txtBox(0) & "', N'" & .Fields("TypeID") & "', " & _
                                   CDbl(.Fields("capacity_qty")) & ",'" & .Fields("capacity_uom") & "','" & _
                                   .Fields("capacity_time_uom") & "','" & .Fields("resource_type") & "')")
                 .MoveNext
            Loop
            End If
            .MoveLast
         End If
    End With
End Sub

Private Sub SavingShift()
Dim cloop As Boolean
    With RcShift.DBRecordset
         Debug.Print .Recordcount
         If .Recordcount <> 0 Then
            .MoveFirst
            If SendDataToServer("Delete From [wcenter_shift] where WCID=N'" & txtBox(0) & "'") = True Then
            cloop = True
            While cloop
                 If .Recordcount = .Bookmark Then cloop = False
                 SendDataToServer (" INSERT INTO [wcenter_shift]" & _
                                   " (ID,WCID, shiftID, shift_desc,start_time,stop_time,break_time,mon_days,tue_days,wed_days,thu_days,fri_days,sat_days,sun_days)" & _
                                   " VALUES (newID(),N'" & txtBox(0) & "', " & .Fields("ShiftID") & ", '" & _
                                   .Fields("shift_desc") & "','" & Format(.Fields("start_time"), "yyyy-MM-dd hh:mm:ss") & "','" & _
                                   Format(.Fields("stop_time"), "yyyy-MM-dd hh:mm:ss") & "'," & CDbl(.Fields("break_time")) & _
                                   "," & IIf(.Fields("mon_days") = True, "1", "0") & "," & IIf(.Fields("tue_days") = True, "1", "0") & "," & IIf(.Fields("wed_days") = True, "1", "0") & _
                                   "," & IIf(.Fields("thu_days") = True, "1", "0") & "," & IIf(.Fields("fri_days") = True, "1", "0") & "," & IIf(.Fields("sat_days") = True, "1", "0") & _
                                   "," & IIf(.Fields("sun_days") = True, "1", "0") & ")")
                 .MoveNext
            Wend
            End If
            .MoveLast
         End If
    End With
End Sub

Private Sub SavingStage()
    With RcStage.DBRecordset
         If .Recordcount <> 0 Then
            .MoveFirst
            If SendDataToServer("Delete From [WC Stage] Where WCID =N'" & txtBox(0) & "'") = True Then
            Do
               If .EOF Then Exit Do
               SendDataToServer (" INSERT INTO [WC Stage]  (WCID, StageID) VALUES (N'" & txtBox(0) & "', N'" & .Fields("StageID") & "')")
               .MoveNext
            Loop
            .MoveLast
            End If
         End If
    End With

End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   OpenDetail MyDDE.GetFieldByName("WCID")
   OpenStage MyDDE.GetFieldByName("WCID")
   Openshift MyDDE.GetFieldByName("WCID")
   OpenBOMWhereUsed MyDDE.GetFieldByName("WCID")
    OpenMOWhereUsed MyDDE.GetFieldByName("WCID")
   CountPerformance
   Me.Caption = "Work Center Setup - " & UCase(MyDDE.GetFieldByName("Description"))
    txtBox(5).Visible = False
End Sub

Private Sub OpenBOMWhereUsed(ByVal Param As String)
    
    RcBOMWhere.DBOpen "SELECT [BOM Stage Detail].ID,[BOM Stage Detail].BomReff,[BOM Stage Detail].WCID,[BOM Stage Detail].Description,[BOM Stage Detail].NoItem,Inventory.ItemName,[BOM Stage Detail].ResourcesID,[BOM Stage Detail].StageNote,[BOM Stage Detail].NoLine From [BOM Stage Detail] INNER JOIN Inventory ON ([BOM Stage Detail].NoItem = Inventory.NoItem) Where [BOM Stage Detail].WCID = '" & Param & "'", CNN, lckLockBatch
    Set GrdBOMWhere.DataSource = RcBOMWhere.DBRecordset
    'Set TDBList1.RowSource = RcBOMWhere.DBRecordset
End Sub
Private Sub OpenMOWhereUsed(ByVal Param As String)
    
    RcMOWhere.DBOpen "SELECT [Order Output Detail].ID,Inventory.ItemName,[Order Output Detail].SeqNo,[Order Output Detail].StartDate,[Order Output Detail].EndDate,[Order Output Detail].Status,[Order Output Detail].WareHouse,[Manufacture Order].OrderName,[Manufacture Order].Type,[Manufacture Order].Status as StatusMO,PartnerDB.CompanyName,[Order Output Detail].ResourcesID,[Resources Table].Description From [Order Output Detail] " & "INNER JOIN [Manufacture Order] ON ([Order Output Detail].OrderID = [Manufacture Order].OrderID) INNER JOIN Inventory ON ([Manufacture Order].NoItem = Inventory.NoItem) INNER JOIN PartnerDB ON (Inventory.PartnerID = PartnerDB.PartnerID) LEFT OUTER JOIN [Resources Table] ON ([Order Output Detail].ResourcesID = [Resources Table].ResourcesID) Where [Order Output Detail].WCID = '" & Param & "'", CNN, lckLockBatch
    Set GrdMOUsed.DataSource = RcMOWhere.DBRecordset
End Sub
Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
   cmdLink(0).Enabled = False

Select Case AdReasonActiveDb
       Case tmbAddNew:
         cmdLink(0).Enabled = True
       Case tmbEdit:
         cmdLink(0).Enabled = True
            If MyDDE.ActiveRecordset.Recordcount <> 0 Then
               If MyDDE.CancelTrans = True Then
                  MessageBox "Transaksi PO Tidak Bisa Diedit.Karena Transaksi PO Sudah Valid/Closed Oleh Transaksi RN."
               End If
            End If
            PrepareSQL
       Case tmbDelete:
            If MyDDE.ActiveRecordset.Recordcount <> 0 Then
               If MyDDE.CancelTrans = True Then
                  MessageBox "Transaksi PO Tidak Bisa Diedit.Karena Transaksi PO Sudah Valid/Closed Oleh Transaksi RN."
               End If
            End If
            PrepareSQL
       Case tmbDetail:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               
            Else
               MyDDE.IsChildMemberReady = False
               MessageBox "Data header transaksi belum lengkap.", "Peringatan"
            End If
       Case tmbSave:
            PrepareSQL
            If MyDDE.CheckEmptyControl = False Then
                  MyDDE.IsChildMemberReady = True
                OpenHeader
            Else
               MyDDE.IsChildMemberReady = False
            End If

         
End Select
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
    PrepareSQL
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MoveForm Picture2.Parent.hwnd
End Sub

Private Sub OptGanttChart_Click()
RepWCInquiry MyDDE.GetFieldByName("WCID")
End Sub

Private Sub OptStackChart_Click()
RepWCInquiry MyDDE.GetFieldByName("WCID")
End Sub
Private Function QuerySource(ByVal vNewValue As String)
    On Error Resume Next
    Rc.DBRecordset.CursorLocation = adUseClient
    Rc.DBOpen vNewValue, CNN, adOpenForwardOnly, adLockReadOnly
    Set Rc.DBRecordset.ActiveConnection = Nothing
    'messagebox Rc.Recordcount
    Err.Clear
End Function

Private Function ReportName(ByVal vNewValue As String)
    On Error GoTo Hell
    Dim mRpt As New CRAXDRT.Report
    Dim mApp As New CRAXDRT.Application
    Set m_Report = m_Application.OpenReport(ReportPath & "\" & vNewValue & ".rpt")
    Exit Function
Hell:


    If Err.Number = -2147189547 Then
        Set mRpt = mApp.OpenReport(ReportPath & "\" & vNewValue & ".rpt")
        Set m_Report = Nothing
        Set m_Report = mRpt
    Else
        MessageBox "PROC_ReportName_ERROR" & vbCrLf & vNewValue & vbCrLf & Err.Description, vbCritical, "Report Warning"
    End If

    Err.Clear
End Function

Private Sub RepWCInquiry(Param As String)
    On Error Resume Next

    If OptGanttChart.Value = True Then
        QuerySource "SELECT * FROM  WC_Inquiry WHERE  WC_Inquiry.WCID ='" & Param & "'"
        ReportName "WC_Inquiry"
    Else
        QuerySource "SELECT * From  WC_ScheduleChart Where WC_ScheduleChart.WCID = '" & Param & "'"
        ReportName "WC_StackedChart"
    End If

    If Not Rc Is Nothing Then
        If Rc.DBRecordset.State = 1 Then
            If Rc.DBRecordset.Recordcount <> 0 Then
                pctReport.Visible = False
                m_Report.DiscardSavedData
                m_Report.Database.SetDataSource Rc
                
                m_Report.ReportComments = GetSetting(App.EXEName, "Lisence Profile", "Address") & vbCrLf & "Telp " & GetSetting(App.EXEName, "Lisence Profile", "Phone") & vbCrLf & GetSetting(App.EXEName, "Lisence Profile", "City")

                If m_Report.ReportAuthor = "" Then m_Report.ReportAuthor = GetSetting(App.EXEName, "Lisence Profile", "Company Name")
                m_Report.EnableParameterPrompting = False
                Me.Caption = m_Report.ReportTitle
           
                CrWCInQuiry.DisplayGroupTree = False
                CrWCInQuiry.ReportSource = m_Report
                CrWCInQuiry.ViewReport
                CrWCInQuiry.Zoom 75
            Else
                pctReport.Visible = True
            End If
        End If
    End If

End Sub

Private Sub RcShift_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo xError
  If RcShift.DBRecordset.Recordcount > 0 Then
    ChkDays(0).Value = IIf(RcShift.DBRecordset.Fields("mon_days") = True, 1, 0)
    ChkDays(1).Value = IIf(RcShift.DBRecordset.Fields("tue_days") = True, 1, 0)
    ChkDays(2).Value = IIf(RcShift.DBRecordset.Fields("wed_days") = True, 1, 0)
    ChkDays(3).Value = IIf(RcShift.DBRecordset.Fields("thu_days") = True, 1, 0)
    ChkDays(4).Value = IIf(RcShift.DBRecordset.Fields("fri_days") = True, 1, 0)
    ChkDays(5).Value = IIf(RcShift.DBRecordset.Fields("sat_days") = True, 1, 0)
    ChkDays(6).Value = IIf(RcShift.DBRecordset.Fields("sun_days") = True, 1, 0)
  End If
Exit Sub
xError:
   If Err.Number = 3021 Then
      Err.Clear
      RcShift.DBRecordset.MoveLast
   Else
      MessageBox Err.Description & " " & Err.Number, , , msgExclamation
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Select Case SSTab1.Tab
           Case 2: Set MyDDE.ChildRecordset = RcShift.DBRecordset
           Case 3: Set MyDDE.ChildRecordset = RcStage.DBRecordset
           Case 4: Set MyDDE.ChildRecordset = RcDetail.DBRecordset
    End Select
End Sub


Private Sub TimeStart_LostFocus(Index As Integer)
    If Index = 7 Then
        Select Case colActive
            Case 4
                RcShift.DBRecordset.Fields("start_time") = TimeStart(7).Value
            Case 5
                RcShift.DBRecordset.Fields("stop_time") = TimeStart(7).Value
        End Select
        TimeStart(7).Visible = False
    End If
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub

Private Sub OpenCalendar()
    RcCal.DBOpen "SELECT CalendarID as Calendar, Description FROM  [Scheduling Calendar]", CNN, lckLockReadOnly
    DataCombo1.DataField = "Calendar"
    DataCombo1.ListField = "Description"
    Set DataCombo1.RowSource = RcCal.DBRecordset
End Sub

Private Sub OpenDetail(ByVal Param As String)
    RcDetail.DBOpen "SELECT [wcenter_resources].TypeID ,[wcenter_resources].TypeID as resources, [Resources Type].Description AS Keterangan, " & _
        " wcenter_resources.capacity_qty, wcenter_resources.capacity_uom, wcenter_resources.capacity_time_uom,wcenter_resources.resource_type,[resources table].description as ResourceName FROM [wcenter_resources] INNER JOIN  " & _
        " [Resources Type] ON [wcenter_resources].TypeID = [Resources Type].TypeID  left outer join [Resources Table] on [Resources Table].ResourcesID = wcenter_resources.resource_type " & _
        " WHERE ([wcenter_resources].WCID = N'" & Param & "') ORDER BY [wcenter_resources].TypeID", CNN, lckLockBatch
'    Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
    Set DataGrid1(1).DataSource = RcDetail.DBRecordset
    GridLayout
End Sub

Private Sub OpenStage(ByVal Param As String)
    
    RcStage.DBOpen "SELECT [WC Stage].WCID,[WC Stage].StageID,[WC Stage].no, [Manufacture Stage].Description FROM [WC Stage] INNER JOIN [Manufacture Stage] ON [WC Stage].StageID = [Manufacture Stage].StageID WHERE ([WC Stage].WCID = N'" & Param & "') ORDER BY [WC Stage].[no]", CNN, lckLockBatch
    'Debug.Print "SELECT [WC Stage].WCID,[WC Stage].StageID,[WC Stage].no, [Manufacture Stage].Description FROM [WC Stage] INNER JOIN [Manufacture Stage] ON [WC Stage].StageID = [Manufacture Stage].StageID WHERE ([WC Stage].WCID = N'" & Param & "') ORDER BY [WC Stage].[no]"
'    Set MyDDE.ChildRecordset = RcStage.DBRecordset.Clone(adLockBatchOptimistic)
    Set DataGrid1(0).DataSource = RcStage.DBRecordset
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:
Select Case Index
    Case 0:
       RcPartner.DBOpen "select formID,formName as [Nama Form] from labformrekomendasi order by FormName", CNN

    Case 1:
        RcPartner.DBOpen " SELECT TypeID AS Resources, Description AS Keterangan FROM         [Resources Type] ORDER BY TypeID", CNN, lckLockReadOnly
        mFirstCaller = True
    Case 2:
        RcPartner.DBOpen " SELECT     StageID, Description  FROM         [Manufacture Stage] ORDER BY StageID", CNN, lckLockReadOnly
        mFirstCaller = True
    Case 3:
        RcPartner.DBOpen " SELECT * FROM [UOM Table] ORDER BY UOM", CNN, lckLockReadOnly
        mFirstCaller = True
    Case 4:
        RcPartner.DBOpen " SELECT Code, Description FROM Capacity_Unit_of_Measure ORDER BY Code", CNN, lckLockReadOnly
        mFirstCaller = True
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 0: mCall.FromTagActive = "Data Form"
          Case 1: mCall.FromTagActive = "DATA RESOURCES"
          Case 2: mCall.FromTagActive = "Data Stage"
          Case 3: mCall.FromTagActive = "Capacity Unit"
          Case 4: mCall.FromTagActive = "Capacity Time Unit"
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me
Else
   MessageBox "Data Belum Ada.", "Peringatan", msgOkOnly
   MyDDE.ChildRecordset.Delete
End If
Exit Sub
Hell:
   MessageBox Err.Description
    Err.Clear
End Sub

Private Sub GridLayout()
'RESOURCES
DataGrid1(1).Columns(0).width = 2025.071
DataGrid1(1).Columns(1).width = 3000
DataGrid1(1).Columns(2).width = 1300
DataGrid1(1).Columns(3).width = 1300
DataGrid1(1).Columns(4).width = 1300
DataGrid1(1).Columns(2).Alignment = dbgRight
DataGrid1(1).Columns(3).Alignment = dbgRight
DataGrid1(1).Columns(4).Alignment = dbgRight
DataGrid1(1).Columns(0).Visible = False
DataGrid1(1).Columns(1).Visible = False

'ROUTINGS / STAGES
DataGrid1(0).Columns(0).width = 2280.189
DataGrid1(0).Columns(1).width = 7680.189
DataGrid1(0).Columns(0).Visible = False
DataGrid1(0).Columns(1).Visible = False

'SHIFT
DataGrid1(2).Columns(0).Visible = False
DataGrid1(2).Columns(1).Visible = False
DataGrid1(2).Columns(7).Visible = False
DataGrid1(2).Columns(8).Visible = False
DataGrid1(2).Columns(9).Visible = False
DataGrid1(2).Columns(10).Visible = False
DataGrid1(2).Columns(11).Visible = False
DataGrid1(2).Columns(12).Visible = False
DataGrid1(2).Columns(13).Visible = False

End Sub


