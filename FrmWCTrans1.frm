VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D7BB8F75-AC9E-4E80-A526-70EA20ACFD16}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmWCTrans 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Work Center  Transaction"
   ClientHeight    =   5835
   ClientLeft      =   3585
   ClientTop       =   3060
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
   Icon            =   "FrmWCTrans1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   Tag             =   "Work Centers"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   5250
      Left            =   0
      ScaleHeight     =   5250
      ScaleWidth      =   10965
      TabIndex        =   1
      Top             =   0
      Width           =   10965
      Begin VB.TextBox txtBox 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DataField       =   "description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   8445
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Tag             =   "Partner"
         Text            =   "Text1"
         Top             =   75
         Width           =   2385
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5070
         Left            =   120
         TabIndex        =   3
         Top             =   105
         Width           =   10770
         _ExtentX        =   18997
         _ExtentY        =   8943
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         TabsPerRow      =   5
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
         TabPicture(0)   =   "FrmWCTrans1.frx":6852
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Picture5"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Detail"
         TabPicture(1)   =   "FrmWCTrans1.frx":686E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Picture3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Shift"
         TabPicture(2)   =   "FrmWCTrans1.frx":688A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Picture6"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Routings"
         TabPicture(3)   =   "FrmWCTrans1.frx":68A6
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Picture4(1)"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Resources"
         TabPicture(4)   =   "FrmWCTrans1.frx":68C2
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Picture4(0)"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).ControlCount=   1
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00EAAF6F&
            Height          =   4545
            Left            =   -74900
            ScaleHeight     =   4485
            ScaleWidth      =   10500
            TabIndex        =   30
            Top             =   400
            Width           =   10560
            Begin VB.Frame Frame1 
               BackColor       =   &H00EAAF6F&
               Caption         =   " Movement "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2850
               Left            =   105
               TabIndex        =   50
               Top             =   1410
               Width           =   4620
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  DataField       =   "queue_time"
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
                  Left            =   2250
                  MaxLength       =   50
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   55
                  Tag             =   "Partner"
                  Top             =   1305
                  Width           =   1140
               End
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  DataField       =   "Utilization"
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
                  Index           =   4
                  Left            =   2250
                  MaxLength       =   50
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   54
                  Tag             =   "Partner"
                  Top             =   2325
                  Width           =   1140
               End
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  DataField       =   "queue_time"
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
                  Index           =   23
                  Left            =   2250
                  MaxLength       =   50
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   53
                  Tag             =   "Partner"
                  Top             =   810
                  Width           =   1140
               End
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  DataField       =   "Utilization"
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
                  Index           =   24
                  Left            =   2250
                  MaxLength       =   50
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   52
                  Tag             =   "Partner"
                  Top             =   1815
                  Width           =   1140
               End
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  DataField       =   "cycle_time"
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
                  Left            =   2250
                  MaxLength       =   50
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   51
                  Tag             =   "Partner"
                  Top             =   300
                  Width           =   1140
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cycle Time (Sec)"
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
                  Index           =   3
                  Left            =   615
                  TabIndex        =   60
                  Top             =   360
                  Width           =   1185
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Queue Time (Sec)"
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
                  Index           =   4
                  Left            =   615
                  TabIndex        =   59
                  Top             =   1372
                  Width           =   1275
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Utilization ( % )"
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
                  Index           =   5
                  Left            =   615
                  TabIndex        =   58
                  Top             =   2385
                  Width           =   1110
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Setup Time (Sec)"
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
                  Index           =   26
                  Left            =   615
                  TabIndex        =   57
                  Top             =   866
                  Width           =   1215
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Wait Time (Sec)"
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
                  Index           =   27
                  Left            =   615
                  TabIndex        =   56
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
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   1440
               MaxLength       =   50
               ScrollBars      =   2  'Vertical
               TabIndex        =   49
               Tag             =   "Partner"
               Top             =   435
               Width           =   3945
            End
            Begin VB.TextBox txtBox 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "Concurrent_Capacities"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   6
               Left            =   7215
               MaxLength       =   10
               TabIndex        =   48
               Tag             =   "Partner"
               Top             =   75
               Width           =   1545
            End
            Begin VB.CommandButton cmdLink 
               Enabled         =   0   'False
               Height          =   315
               Index           =   3
               Left            =   8760
               Picture         =   "FrmWCTrans1.frx":68DE
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   443
               Width           =   345
            End
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "Capacity_UOM"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   12
               Left            =   7215
               MaxLength       =   15
               TabIndex        =   46
               Tag             =   "Partner"
               Top             =   435
               Width           =   1545
            End
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "Capacity_Time_UOM_Code"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   7
               Left            =   7215
               MaxLength       =   10
               TabIndex        =   45
               Tag             =   "Partner"
               Top             =   795
               Width           =   1545
            End
            Begin VB.CommandButton cmdLink 
               Enabled         =   0   'False
               Height          =   315
               Index           =   4
               Left            =   8760
               Picture         =   "FrmWCTrans1.frx":6C68
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   803
               Width           =   345
            End
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "WCID"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   1440
               MaxLength       =   10
               TabIndex        =   43
               Tag             =   "Partner"
               Top             =   75
               Width           =   3945
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00EAAF6F&
               Caption         =   " Performance "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2235
               Left            =   4920
               TabIndex        =   31
               Top             =   1410
               Width           =   5445
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  DataField       =   "Utilization"
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
                  Index           =   17
                  Left            =   1860
                  MaxLength       =   50
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   67
                  Tag             =   "Partner"
                  Top             =   1710
                  Width           =   1110
               End
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  DataField       =   "Queue Time"
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
                  Index           =   19
                  Left            =   1860
                  MaxLength       =   50
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   36
                  Tag             =   "Partner"
                  Top             =   690
                  Width           =   1110
               End
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  DataField       =   "Move Time"
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
                  Index           =   18
                  Left            =   1860
                  MaxLength       =   50
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   35
                  Tag             =   "Partner"
                  Top             =   1200
                  Width           =   1110
               End
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  DataField       =   "Utilization"
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
                  Index           =   20
                  Left            =   4140
                  MaxLength       =   50
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   34
                  Tag             =   "Partner"
                  Top             =   1710
                  Width           =   1110
               End
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  DataField       =   "Move Time"
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
                  Index           =   21
                  Left            =   4140
                  MaxLength       =   50
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   33
                  Tag             =   "Partner"
                  Top             =   1200
                  Width           =   1110
               End
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  DataField       =   "Queue Time"
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
                  Index           =   22
                  Left            =   4140
                  MaxLength       =   50
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   32
                  Tag             =   "Partner"
                  Top             =   690
                  Width           =   1110
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Overhead"
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
                  Index           =   6
                  Left            =   285
                  TabIndex        =   41
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
                  Index           =   7
                  Left            =   285
                  TabIndex        =   40
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
                  Index           =   8
                  Left            =   285
                  TabIndex        =   39
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
                  Index           =   24
                  Left            =   2085
                  TabIndex        =   38
                  Top             =   315
                  Width           =   630
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Per Unit"
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
                  Index           =   25
                  Left            =   4395
                  TabIndex        =   37
                  Top             =   330
                  Width           =   570
               End
            End
            Begin MSDataListLib.DataCombo DataCombo1 
               DataField       =   "Calendar"
               Height          =   315
               Left            =   1440
               TabIndex        =   42
               Tag             =   "Partner"
               Top             =   795
               Width           =   3945
               _ExtentX        =   6959
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               ListField       =   "Description"
               BoundColumn     =   "Calendar"
               Text            =   "DataCombo1"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
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
               Index           =   0
               Left            =   195
               TabIndex        =   66
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
               Index           =   2
               Left            =   195
               TabIndex        =   65
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
               Caption         =   "Concurrent Capacity"
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
               Index           =   9
               Left            =   5580
               TabIndex        =   64
               Top             =   150
               Width           =   1485
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Capacity UOM"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   22
               Left            =   5580
               TabIndex        =   63
               Top             =   503
               Width           =   1020
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
               Index           =   1
               Left            =   195
               TabIndex        =   62
               Top             =   143
               Width           =   915
            End
            Begin VB.Line Line1 
               Index           =   7
               X1              =   5565
               X2              =   7230
               Y1              =   1110
               Y2              =   1110
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Capacity TIME"
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
               Index           =   10
               Left            =   5580
               TabIndex        =   61
               Top             =   870
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
            TabIndex        =   28
            Top             =   400
            Width           =   10560
            Begin MSComctlLib.ListView ListView1 
               Height          =   4500
               Left            =   0
               TabIndex        =   29
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
            TabIndex        =   26
            Top             =   400
            Width           =   10560
            Begin MSDataGridLib.DataGrid DataGrid1 
               Bindings        =   "FrmWCTrans1.frx":6FF2
               Height          =   4515
               Index           =   0
               Left            =   -15
               TabIndex        =   27
               Tag             =   "partner"
               Top             =   -15
               Width           =   10530
               _ExtentX        =   18574
               _ExtentY        =   7964
               _Version        =   393216
               AllowUpdate     =   -1  'True
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
                  DataField       =   "No"
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
                     ColumnWidth     =   1739,906
                  EndProperty
                  BeginProperty Column01 
                     ColumnWidth     =   1140,095
                  EndProperty
                  BeginProperty Column02 
                  EndProperty
                  BeginProperty Column03 
                     ColumnWidth     =   915,024
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
            TabIndex        =   24
            Top             =   400
            Width           =   10560
            Begin MSDataGridLib.DataGrid DataGrid1 
               Bindings        =   "FrmWCTrans1.frx":7007
               Height          =   4515
               Index           =   1
               Left            =   -15
               TabIndex        =   25
               Tag             =   "partner"
               Top             =   -30
               Width           =   10545
               _ExtentX        =   18600
               _ExtentY        =   7964
               _Version        =   393216
               AllowUpdate     =   -1  'True
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
               ColumnCount     =   6
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
               SplitCount      =   1
               BeginProperty Split0 
                  BeginProperty Column00 
                     ColumnWidth     =   1739,906
                  EndProperty
                  BeginProperty Column01 
                     ColumnWidth     =   1140,095
                  EndProperty
                  BeginProperty Column02 
                     ColumnWidth     =   1140,095
                  EndProperty
                  BeginProperty Column03 
                     ColumnWidth     =   1739,906
                  EndProperty
                  BeginProperty Column04 
                     ColumnWidth     =   1649,764
                  EndProperty
                  BeginProperty Column05 
                     ColumnWidth     =   2564,788
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
            TabIndex        =   4
            Top             =   405
            Width           =   10560
            Begin VB.Frame Frame3 
               BackColor       =   &H00EAAF6F&
               Caption         =   " Jadwal Kerja "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1620
               Left            =   90
               TabIndex        =   6
               Top             =   2520
               Width           =   10290
               Begin VB.CheckBox ChkDays 
                  BackColor       =   &H00EAAF6F&
                  DataSource      =   "MyDDE"
                  Height          =   225
                  Index           =   6
                  Left            =   8310
                  TabIndex        =   13
                  Top             =   915
                  Width           =   255
               End
               Begin VB.CheckBox ChkDays 
                  BackColor       =   &H00EAAF6F&
                  DataSource      =   "MyDDE"
                  Height          =   225
                  Index           =   5
                  Left            =   7198
                  TabIndex        =   12
                  Top             =   915
                  Width           =   255
               End
               Begin VB.CheckBox ChkDays 
                  BackColor       =   &H00EAAF6F&
                  DataSource      =   "MyDDE"
                  Height          =   225
                  Index           =   4
                  Left            =   6087
                  TabIndex        =   11
                  Top             =   915
                  Width           =   255
               End
               Begin VB.CheckBox ChkDays 
                  BackColor       =   &H00EAAF6F&
                  DataSource      =   "MyDDE"
                  Height          =   225
                  Index           =   3
                  Left            =   4976
                  TabIndex        =   10
                  Top             =   915
                  Width           =   255
               End
               Begin VB.CheckBox ChkDays 
                  BackColor       =   &H00EAAF6F&
                  DataSource      =   "MyDDE"
                  Height          =   225
                  Index           =   2
                  Left            =   3865
                  TabIndex        =   9
                  Top             =   915
                  Width           =   255
               End
               Begin VB.CheckBox ChkDays 
                  BackColor       =   &H00EAAF6F&
                  DataSource      =   "MyDDE"
                  Height          =   225
                  Index           =   1
                  Left            =   2754
                  TabIndex        =   8
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
                  TabIndex        =   7
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
                  TabIndex        =   22
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
                  TabIndex        =   21
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
                  TabIndex        =   20
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
                  TabIndex        =   19
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
                  TabIndex        =   18
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
                  TabIndex        =   17
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
                  TabIndex        =   16
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
                  TabIndex        =   15
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
                  TabIndex        =   14
                  Top             =   885
                  Width           =   525
               End
            End
            Begin MSComCtl2.DTPicker TimeStart 
               DataSource      =   "MyDDE"
               Height          =   315
               Index           =   7
               Left            =   2745
               TabIndex        =   5
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
               Format          =   17563650
               UpDown          =   -1  'True
               CurrentDate     =   38467
            End
            Begin MSDataGridLib.DataGrid DataGrid1 
               Bindings        =   "FrmWCTrans1.frx":701C
               Height          =   2385
               Index           =   2
               Left            =   -15
               TabIndex        =   23
               Tag             =   "partner"
               Top             =   -15
               Width           =   10515
               _ExtentX        =   18547
               _ExtentY        =   4207
               _Version        =   393216
               AllowUpdate     =   -1  'True
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
                     ColumnWidth     =   1739,906
                  EndProperty
                  BeginProperty Column01 
                     ColumnWidth     =   1140,095
                  EndProperty
                  BeginProperty Column02 
                     ColumnWidth     =   1140,095
                  EndProperty
                  BeginProperty Column03 
                     ColumnWidth     =   1140,095
                  EndProperty
                  BeginProperty Column04 
                     ColumnWidth     =   1739,906
                  EndProperty
                  BeginProperty Column05 
                     ColumnWidth     =   1739,906
                  EndProperty
                  BeginProperty Column06 
                     ColumnWidth     =   1739,906
                  EndProperty
                  BeginProperty Column07 
                     ColumnWidth     =   900,284
                  EndProperty
                  BeginProperty Column08 
                     ColumnWidth     =   824,882
                  EndProperty
                  BeginProperty Column09 
                     ColumnWidth     =   900,284
                  EndProperty
                  BeginProperty Column10 
                     ColumnWidth     =   824,882
                  EndProperty
                  BeginProperty Column11 
                     ColumnWidth     =   764,787
                  EndProperty
                  BeginProperty Column12 
                     ColumnWidth     =   780,095
                  EndProperty
                  BeginProperty Column13 
                     ColumnWidth     =   824,882
                  EndProperty
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
      Top             =   5265
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmWCTrans"
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

Private Sub ChkDays_Click(Index As Integer)

    Select Case Index

        Case 0: RcShift.DBRecordset.Fields("mon_days") = ChkDays(Index).value

        Case 1: RcShift.DBRecordset.Fields("tue_days") = ChkDays(Index).value

        Case 2: RcShift.DBRecordset.Fields("wed_days") = ChkDays(Index).value

        Case 3: RcShift.DBRecordset.Fields("thu_days") = ChkDays(Index).value

        Case 4: RcShift.DBRecordset.Fields("fri_days") = ChkDays(Index).value

        Case 5: RcShift.DBRecordset.Fields("sat_days") = ChkDays(Index).value

        Case 6: RcShift.DBRecordset.Fields("sun_days") = ChkDays(Index).value
    End Select

End Sub

Private Sub cmdLink_Click(Index As Integer)
    OpenPartner Index
End Sub

Private Sub DataCombo1_KeyDown(KeyCode As Integer, _
                               Shift As Integer)
    KeyEnter KeyCode
End Sub

Private Sub DataGrid1_BeforeColEdit(Index As Integer, _
                                    ByVal ColIndex As Integer, _
                                    ByVal KeyAscii As Integer, _
                                    Cancel As Integer)

    Select Case Index

        Case 2:

            If (ColIndex = 4) Or (ColIndex = 5) Then
                TimeStart(7).Visible = True
                TimeStart(7).Move DataGrid1(Index).Columns(ColIndex).Left, DataGrid1(Index).RowTop(DataGrid1(Index).Row), DataGrid1(Index).Columns(ColIndex).Width, DataGrid1(Index).RowHeight
                TimeStart(7).SetFocus
                
            End If

    End Select

    colActive = ColIndex
End Sub

Private Sub Form_Load()
    Set RcShift = New DBQuick
    HiasFormManTell Picture2, Me
    SSTab1.tab = 0
    OpenCalendar

    With MyDDE
        .EditModeReplace = False
        Set .BindForm = FrmWCTrans
        .BindFormTAG = "Partner"
        Set .ActiveConnection = CNN
        .PrepareQuery = "SELECT WCID, Description AS Keterangan, CalendarID AS Calendar, " & " cycle_time, queue_time, setup_time, wait_time, utilization, " & " Concurrent_Capacities, Capacity_UOM, Capacity_Time_UOM_Code,Description FROM wcenter_header ORDER BY WCID"
    End With

    OpenHeader
    OpenChild
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

Private Sub prepareSQL()

    With MyDDE
        .PrepareAppend = "insert into wcenter_header ([WCID] ,[CalendarID],[Description],[cycle_time],[queue_time]," & _
           " [setup_time],[wait_time],[utilization],[Concurrent_Capacities],[Maximum_Efficiency]," & _
           " [Minimum_Efficiency],[Blocked],[Fixed_Scrap_Quantity],[Scrap],[Capacity_UOM],[Capacity_Time_UOM_Code]," & _
           " [cycle_time_uom_code],[Queue_Time_UOM_Code],[Setup_Time_UOM_Code],[Wait_Time_UOM_Code]," & _
           " [Flushing_Method],[No_Series],[Overhead_Rate],[Gen_Prod_Posting_Group]) values ('" & _
           .GetFieldByName("WCID") & "','" & DataCombo1.BoundText & "','" & .GetFieldByName("Description") & _
           "'," & FQty(.GetFieldByName("cycle_time")) & "," & FQty(.GetFieldByName("queue_time")) & "," & _
           FQty(.GetFieldByName("setup_time")) & "," & FQty(.GetFieldByName("wait_time")) & "," & FQty(.GetFieldByName("utilization")) & _
           "," & FQty(.GetFieldByName("conCurrent_capacities")) & "," & FQty(.GetFieldByName("Maximum_Efficiency")) & "," & FQty(.GetFieldByName("Maximum_Efficiency")) & _
           "," & FQty(.GetFieldByName("Blocked")) & "," & FQty(.GetFieldByName("Fixed_Scrap_Quantity")) & "," & _
           FQty(.GetFieldByName("scrap")) & ",'" & .GetFieldByName("capacity_UOM") & "','" & .GetFieldByName("capacity_Time_UOM_code") & _
           "','" & .GetFieldByName("cycle_time_uom_code") & "','" & .GetFieldByName("queue_time_uom_code") & "','" & _
           .GetFieldByName("setup_time_uom_code") & "','" & .GetFieldByName("wait_time_UOM_code") & "'," & _
           FQty(.GetFieldByName("flushing_method")) & ",'" & .GetFieldByName("no_series") & "'," & FQty(.GetFieldByName("Overhead_rate")) & _
           ",'" & .GetFieldByName("gen_prod_posting_group") & "')"
   
        .PrepareUpdate = " update wcenter_header set [CalendarID] ='" & DataCombo1.BoundText & "'," & _
           "[Description]='" & .GetFieldByName("Description") & "'," & _
           "[cycle_time] = " & FQty(.GetFieldByName("cycle_time")) & "," & _
           "[queue_time] = " & FQty(.GetFieldByName("queue_time")) & "," & _
           "[setup_time] = " & FQty(.GetFieldByName("setup_time")) & "," & _
           "[wait_time]  = " & FQty(.GetFieldByName("wait_time")) & "," & _
           "[utilization]= " & FQty(.GetFieldByName("utilization")) & "," & _
           "[Concurrent_Capacities] = " & FQty(.GetFieldByName("conCurrent_capacities")) & "," & _
           "[Maximum_Efficiency] = " & FQty(.GetFieldByName("Maximum_Efficiency")) & "," & _
           "[Minimum_Efficiency] = " & FQty(.GetFieldByName("Minimum_Efficiency")) & "," & _
           "[Blocked] = " & FQty(.GetFieldByName("Blocked")) & "," & _
           "[Fixed_Scrap_Quantity] = " & FQty(.GetFieldByName("Fixed_Scrap_Quantity")) & "," & _
           "[Scrap] = " & FQty(.GetFieldByName("scrap")) & "," & _
           "[Capacity_UOM] = '" & .GetFieldByName("capacity_UOM") & "'," & _
           "[Capacity_Time_UOM_Code] = '" & .GetFieldByName("capacity_Time_UOM_code") & "'," & _
           "[cycle_time_uom_code] = '" & .GetFieldByName("cycle_time_uom_code") & "'," & _
           "[Queue_Time_UOM_Code] = '" & .GetFieldByName("queue_time_uom_code") & "'," & _
           "[Setup_Time_UOM_Code] = '" & .GetFieldByName("setup_time_uom_code") & "'," & _
           "[Wait_Time_UOM_Code] = '" & .GetFieldByName("wait_time_UOM_code") & "'," & _
           "[Flushing_Method] = " & FQty(.GetFieldByName("flushing_method")) & "," & _
           "[No_Series] = '" & .GetFieldByName("no_series") & "'," & _
           "[Overhead_Rate] = " & FQty(.GetFieldByName("Overhead_rate")) & "," & _
           "[Gen_Prod_Posting_Group] ='" & .GetFieldByName("gen_prod_posting_group") & "'" & _
           " where WCID ='" & .GetFieldByName("WCID") & "'"
                     
        .PrepareDelete = " delete from wcenter_header where WCID ='" & .GetFieldByName("WCID") & "'"
                     
    End With

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    ScanKey KeyCode, Shift, MyDDE
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

Private Sub ListView1_DblClick()
    SSTab1.tab = 1
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If MyDDE.ActiveRecordset.Recordcount <> 0 Then
        MyDDE.FindStringData "[WCID]='" & Item.Text & "'"
    End If

End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, _
                               ByVal pRecordset As ADODB.Recordset)

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
                .GetFieldByName("Capacity_UOM") = mCall.GetFieldByName(0)
                mFirstCaller = False
            End With
            
        Case "CAPACITY TIME UNIT":

            With MyDDE
                .GetFieldByName("Capacity_Time_UOM_Code") = mCall.GetFieldByName(0)
                mFirstCaller = False
            End With

    End Select

End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
    On Error GoTo AfterErr

    Select Case AdReasonActiveDb

        Case tmbEdit:
            MEdit = True

            If SSTab1.tab = 0 Then SSTab1.tab = 1
            txtBox(0).Enabled = False
            txtBox(1).SetFocus
            DataGrid1(0).AllowUpdate = True

        Case tmbAddNew:
            MEdit = True
            MyDDE.GetFieldByName("WCID") = "New Order"
            MyDDE.GetFieldByName("Keterangan") = "-"
            SSTab1.tab = 1
            txtBox(0).Enabled = True
            txtBox(1).SetFocus
            DataGrid1(0).AllowUpdate = True

        Case tmbSave:

            If MyDDE.IsChildMemberReady = True Then
                SavingResource
                SavingShift
                SavingStage
                MEdit = False
            End If

        Case tmbCancel:

        Case tmbDetail:

            If mFirstCaller = False Then

                Select Case SSTab1.tab

                    Case 4: OpenPartner 1

                    Case 3: OpenPartner 2
                End Select

                MEdit = True
            End If

        Case tmbPrint:
            CallRPTReport "Manufacture WC Table.rpt", "Select * From [Manufacture WC Table] where [Work ID] ='" & txtBox(0) & "'"

        Case tmbQuit:
            Unload Me
            Set MyDDE.BindForm = Nothing
    End Select

    CmdLink(3).Enabled = MEdit
    CmdLink(4).Enabled = MEdit
    Exit Sub
AfterErr:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub SavingResource()

    With RcDetail.DBRecordset

        If .Recordcount <> 0 Then
            .MoveFirst

            If SendDataToServer("Delete From [wcenter_resources] where WCID=N'" & txtBox(0) & "'") = True Then

                Do

                    If .EOF Then Exit Do
                    SendDataToServer (" INSERT INTO [wcenter_resources]" & " (WCID, TypeID, capacity_qty,capacity_uom,capacity_time_uom)" & " VALUES (N'" & txtBox(0) & "', N'" & .Fields("TypeID") & "', " & CDbl(.Fields("capacity_qty")) & ",'" & .Fields("capacity_uom") & "','" & .Fields("capacity_time_uom") & "')")
                    .MoveNext
                Loop

            End If

            .MoveLast
        End If

    End With

End Sub

Private Sub SavingShift()

    With RcDetail.DBRecordset

        If .Recordcount <> 0 Then
            .MoveFirst

            If SendDataToServer("Delete From [wcenter_shift] where WCID=N'" & txtBox(0) & "'") = True Then

                Do

                    If .EOF Then Exit Do
                    SendDataToServer (" INSERT INTO [wcenter_shift]" & " (WCID, shiftID, shift_desc,start_time,stop_time,break_time,mon_days,tue_days,wed_days,thu_days,fri_days,sat_days,sun_days)" & " VALUES (N'" & txtBox(0) & "', " & .Fields("shiftID") & ", '" & .Fields("shift_desc") & "','" & Format(.Fields("start_time"), "yyyy-MM-dd hh:mm:ss") & "','" & Format(.Fields("stop_time"), "yyyy-MM-dd hh:mm:ss") & "'," & CDbl(.Fields("break_time")) & "," & .Fields("mon_days") & "," & .Fields("tue_days") & "," & .Fields("wed_days") & "," & .Fields("thu_days") & "," & .Fields("fri_days") & "," & .Fields("sat_days") & "," & .Fields("sun_days") & ")")
                    .MoveNext
                Loop

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
                    SendDataToServer (" INSERT INTO [WC Stage]  (WCID, StageID,no) VALUES (N'" & txtBox(0) & "', N'" & .Fields(0) & "'," & .Fields("no") & ")")
                    .MoveNext
                Loop

                .MoveLast
            End If
        End If

    End With

End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
                               ByVal pError As ADODB.Error, _
                               adStatus As ADODB.EventStatusEnum, _
                               ByVal pRecordset As ADODB.Recordset)
    OpenDetail MyDDE.GetFieldByName("WCID")
    OpenStage MyDDE.GetFieldByName("WCID")
    Openshift MyDDE.GetFieldByName("WCID")
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb

        Case tmbEdit, tmbDelete:

            If MyDDE.ActiveRecordset.Recordcount <> 0 Then
                If MyDDE.CancelTrans = True Then
                    MessageBox "Transaksi PO Tidak Bisa Diedit.Karena Transaksi PO Sudah Valid/Closed Oleh Transaksi RN."
                End If
            End If

        Case tmbDetail:

            If MyDDE.CheckEmptyControl = False Then
                MyDDE.IsChildMemberReady = True
               
            Else
                MyDDE.IsChildMemberReady = False
                MessageBox "Data header transaksi belum lengkap.", "Peringatan"
            End If

        Case tmbSave:
            prepareSQL

            If MyDDE.CheckEmptyControl = False Then
                MyDDE.IsChildMemberReady = True
                OpenHeader
            Else
                MyDDE.IsChildMemberReady = False
            End If

    End Select

End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
    prepareSQL
End Sub

Private Sub Picture1_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    MoveForm Picture2.Parent.hwnd
End Sub

Private Sub RcShift_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
                                 ByVal pError As ADODB.Error, _
                                 adStatus As ADODB.EventStatusEnum, _
                                 ByVal pRecordset As ADODB.Recordset)

    If RcShift.DBRecordset.Recordcount > 0 Then
        ChkDays(0).value = IIf(RcShift.DBRecordset.Fields("mon_days") = True, 1, 0)
        ChkDays(1).value = IIf(RcShift.DBRecordset.Fields("tue_days") = True, 1, 0)
        ChkDays(2).value = IIf(RcShift.DBRecordset.Fields("wed_days") = True, 1, 0)
        ChkDays(3).value = IIf(RcShift.DBRecordset.Fields("thu_days") = True, 1, 0)
        ChkDays(4).value = IIf(RcShift.DBRecordset.Fields("fri_days") = True, 1, 0)
        ChkDays(5).value = IIf(RcShift.DBRecordset.Fields("sat_days") = True, 1, 0)
        ChkDays(6).value = IIf(RcShift.DBRecordset.Fields("sun_days") = True, 1, 0)
    End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

    Select Case SSTab1.tab

        Case 2: Set MyDDE.ChildRecordset = RcShift.DBRecordset

        Case 3: Set MyDDE.ChildRecordset = RcStage.DBRecordset

        Case 4: Set MyDDE.ChildRecordset = RcDetail.DBRecordset
    End Select

End Sub

Private Sub TimeStart_LostFocus(Index As Integer)

    If Index = 7 Then

        Select Case colActive

            Case 4
                RcShift.DBRecordset.Fields("start_time") = TimeStart(7).value

            Case 5
                RcShift.DBRecordset.Fields("stop_time") = TimeStart(7).value
        End Select

        TimeStart(7).Visible = False
    End If

End Sub

Private Sub txtBox_KeyDown(Index As Integer, _
                           KeyCode As Integer, _
                           Shift As Integer)
    KeyEnter KeyCode
End Sub

Private Sub OpenCalendar()
    RcCal.DBOpen "SELECT CalendarID as Calendar, Description FROM  [Scheduling Calendar]", CNN, lckLockReadOnly
    DataCombo1.DataField = "Calendar"
    DataCombo1.ListField = "Description"
    Set DataCombo1.RowSource = RcCal.DBRecordset
End Sub

Private Sub OpenDetail(ByVal Param As String)
    RcDetail.DBOpen "SELECT [wcenter_resources].TypeID ,[wcenter_resources].TypeID as resources, [Resources Type].Description AS Keterangan, " & " wcenter_resources.capacity_qty, wcenter_resources.capacity_uom, wcenter_resources.capacity_time_uom FROM [wcenter_resources] INNER JOIN  " & " [Resources Type] ON [wcenter_resources].TypeID = [Resources Type].TypeID " & " WHERE ([wcenter_resources].WCID = N'" & Param & "') ORDER BY [wcenter_resources].TypeID", CNN, lckLockBatch
    '    Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
    Set DataGrid1(1).DataSource = RcDetail.DBRecordset
    GridLayout
End Sub

Private Sub OpenStage(ByVal Param As String)
    RcStage.DBOpen "SELECT [WC Stage].WCID,[WC Stage].StageID,[WC Stage].no, [Manufacture Stage].Description FROM [WC Stage] INNER JOIN [Manufacture Stage] ON [WC Stage].StageID = [Manufacture Stage].StageID WHERE ([WC Stage].WCID = N'" & Param & "') ORDER BY [WC Stage].[no]", CNN, lckLockBatch
    '    Set MyDDE.ChildRecordset = RcStage.DBRecordset.Clone(adLockBatchOptimistic)
    Set DataGrid1(0).DataSource = RcStage.DBRecordset
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
    On Error GoTo Hell:

    Select Case Index

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

            Case 1: mCall.FromTagActive = "DATA RESOURCES"

            Case 2: mCall.FromTagActive = "Data Stage"

            Case 3: mCall.FromTagActive = "Capacity Unit"

            Case 4: mCall.FromTagActive = "Capacity Time Unit"
        End Select

        Set mCall.FormData = RcPartner.DBRecordset
        mCall.LookUp Me
    Else
        MessageBox "Data Belum Ada.", "Peringatan", msgOkOnly
    End If

    Exit Sub
Hell:
    Err.Clear
End Sub

Private Sub GridLayout()
    'RESOURCES
    DataGrid1(1).Columns(0).Width = 2025.071
    DataGrid1(1).Columns(1).Width = 3000
    DataGrid1(1).Columns(2).Width = 1300
    DataGrid1(1).Columns(3).Width = 1300
    DataGrid1(1).Columns(4).Width = 1300
    DataGrid1(1).Columns(2).Alignment = dbgRight
    DataGrid1(1).Columns(3).Alignment = dbgRight
    DataGrid1(1).Columns(4).Alignment = dbgRight
    DataGrid1(1).Columns(0).Visible = False
    DataGrid1(1).Columns(1).Visible = False

    'ROUTINGS / STAGES
    DataGrid1(0).Columns(0).Width = 2280.189
    DataGrid1(0).Columns(1).Width = 7680.189
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

