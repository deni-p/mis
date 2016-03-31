VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmItemData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item Data"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11460
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmItemData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   Tag             =   "Inventory Card"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8280
      Left            =   0
      ScaleHeight     =   8280
      ScaleWidth      =   11460
      TabIndex        =   48
      Top             =   0
      Width           =   11460
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   5
         Left            =   8445
         Picture         =   "FrmItemData.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2228
         Width           =   330
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   8445
         Picture         =   "FrmItemData.frx":6BDC
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1148
         Width           =   330
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   8445
         Picture         =   "FrmItemData.frx":6F66
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   428
         Width           =   330
      End
      Begin VB.TextBox Label2 
         Appearance      =   0  'Flat
         DataField       =   "CompanyName"
         Height          =   330
         Index           =   2
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   9
         Tag             =   "Partner"
         Top             =   1875
         Width           =   3045
      End
      Begin VB.TextBox Label2 
         Appearance      =   0  'Flat
         DataField       =   "Nama Gudang Persediaan"
         Height          =   330
         Index           =   1
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   7
         Tag             =   "Partner"
         Top             =   1515
         Width           =   3045
      End
      Begin VB.TextBox Label2 
         Appearance      =   0  'Flat
         DataField       =   "Kelompok Persediaan"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   0
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   3
         Tag             =   "Partner"
         Top             =   795
         Width           =   3045
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Manufacture"
         DataSource      =   "MyDDE"
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
         Height          =   270
         Left            =   1725
         TabIndex        =   15
         Tag             =   "Partner"
         Top             =   3690
         Width           =   1920
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Reference"
         Height          =   450
         Index           =   1
         Left            =   9825
         TabIndex        =   35
         Top             =   3900
         Width           =   1485
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Qty Availlable"
         Height          =   450
         Index           =   0
         Left            =   9825
         TabIndex        =   34
         Top             =   3390
         Width           =   1485
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   4785
         Picture         =   "FrmItemData.frx":72F0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1883
         Width           =   330
      End
      Begin VB.TextBox txtBox 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "ROL"
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
         Index           =   7
         Left            =   6630
         MaxLength       =   15
         TabIndex        =   32
         Tag             =   "Partner"
         Text            =   " "
         Top             =   3660
         Width           =   1815
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   4785
         Picture         =   "FrmItemData.frx":767A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1523
         Width           =   330
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   4785
         Picture         =   "FrmItemData.frx":7A04
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   803
         Width           =   330
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "NoItem"
         Height          =   330
         Index           =   0
         Left            =   1740
         MaxLength       =   15
         TabIndex        =   1
         Tag             =   "Partner"
         Top             =   75
         Width           =   1935
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "ItemName"
         Height          =   330
         Index           =   1
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "Partner"
         Top             =   435
         Width           =   3045
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "category"
         Height          =   330
         Index           =   2
         Left            =   1740
         MaxLength       =   25
         TabIndex        =   5
         Tag             =   "Partner"
         Top             =   1155
         Width           =   3045
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "partnerID"
         Height          =   330
         Index           =   3
         Left            =   1740
         MaxLength       =   25
         TabIndex        =   11
         Tag             =   "Partner"
         Top             =   2257
         Width           =   3045
      End
      Begin VB.TextBox txtBox 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "MinStock"
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
         Left            =   6630
         MaxLength       =   15
         TabIndex        =   29
         Tag             =   "Partner"
         Top             =   2580
         Width           =   1815
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Purchase"
         DataSource      =   "MyDDE"
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
         Height          =   270
         Left            =   1725
         TabIndex        =   16
         Tag             =   "Partner"
         Top             =   3945
         Width           =   1920
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Pre-Manufacture"
         DataSource      =   "MyDDE"
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
         Height          =   270
         Left            =   1725
         TabIndex        =   17
         Tag             =   "Partner"
         Top             =   4200
         Width           =   1920
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "InternalName"
         Height          =   330
         Index           =   14
         Left            =   6630
         MaxLength       =   50
         TabIndex        =   18
         Tag             =   "Partner"
         Top             =   60
         Width           =   3045
      End
      Begin VB.TextBox txtBox 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "MaxStock"
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
         Left            =   6630
         MaxLength       =   15
         TabIndex        =   30
         Tag             =   "Partner"
         Text            =   " "
         Top             =   2940
         Width           =   1815
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "UOMPurchase"
         Height          =   330
         Index           =   4
         Left            =   6630
         MaxLength       =   15
         TabIndex        =   19
         Tag             =   "Partner"
         Top             =   420
         Width           =   1815
      End
      Begin VB.TextBox txtBox 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "ROP"
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
         Index           =   8
         Left            =   6630
         MaxLength       =   15
         TabIndex        =   31
         Tag             =   "Partner"
         Text            =   " "
         Top             =   3300
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "Status"
         Height          =   315
         ItemData        =   "FrmItemData.frx":7D8E
         Left            =   1725
         List            =   "FrmItemData.frx":7D9E
         TabIndex        =   12
         Tag             =   "Partner"
         Text            =   "- Status -"
         Top             =   2613
         Width           =   1560
      End
      Begin VB.TextBox txtBox 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "Harga"
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
         Index           =   9
         Left            =   6630
         MaxLength       =   15
         TabIndex        =   26
         Tag             =   "Partner"
         Top             =   1860
         Width           =   1815
      End
      Begin VB.TextBox txtBox 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "LeadTimeDays"
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
         Index           =   10
         Left            =   6630
         MaxLength       =   15
         TabIndex        =   33
         Tag             =   "Partner"
         Text            =   " "
         Top             =   4020
         Width           =   1815
      End
      Begin VB.TextBox txtBox 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "UOMKonversi"
         Height          =   330
         Index           =   11
         Left            =   6630
         MaxLength       =   15
         TabIndex        =   21
         Tag             =   "Partner"
         Top             =   780
         Width           =   1815
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "UOM"
         Height          =   330
         Index           =   12
         Left            =   6630
         MaxLength       =   15
         TabIndex        =   22
         Tag             =   "Partner"
         Top             =   1140
         Width           =   1815
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
         Index           =   13
         Left            =   6630
         MaxLength       =   15
         TabIndex        =   27
         Tag             =   "Partner"
         Top             =   2220
         Width           =   1815
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   6
         Left            =   8445
         Picture         =   "FrmItemData.frx":7DC6
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1508
         Width           =   330
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "UOMSales"
         Height          =   330
         Index           =   15
         Left            =   6615
         MaxLength       =   15
         TabIndex        =   24
         Tag             =   "Partner"
         Top             =   1500
         Width           =   1815
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   7
         Left            =   4770
         Picture         =   "FrmItemData.frx":8150
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "Partner"
         Top             =   2963
         Width           =   330
      End
      Begin VB.TextBox Label2 
         Appearance      =   0  'Flat
         DataField       =   "TrackingName"
         Height          =   330
         Index           =   3
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   13
         Tag             =   "Partner"
         Top             =   2955
         Width           =   3045
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   8
         Left            =   4785
         Picture         =   "FrmItemData.frx":84DA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1163
         Width           =   330
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   2910
         Left            =   105
         TabIndex        =   36
         Top             =   5265
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   5133
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         BackColor       =   15380335
         TabCaption(0)   =   "Inventory"
         TabPicture(0)   =   "FrmItemData.frx":8864
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Picture3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Where Used BOM"
         TabPicture(1)   =   "FrmItemData.frx":8880
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Picture4"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "History"
         TabPicture(2)   =   "FrmItemData.frx":889C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Picture5"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Hazard Identification"
         TabPicture(3)   =   "FrmItemData.frx":88B8
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Picture1"
         Tab(3).ControlCount=   1
         Begin VB.PictureBox Picture1 
            Height          =   2475
            Left            =   -74925
            ScaleHeight     =   2415
            ScaleWidth      =   11055
            TabIndex        =   82
            Top             =   360
            Width           =   11115
            Begin MSDataGridLib.DataGrid GridHazard 
               Height          =   2445
               Left            =   0
               TabIndex        =   83
               Top             =   0
               Width           =   11130
               _ExtentX        =   19632
               _ExtentY        =   4313
               _Version        =   393216
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
               ColumnCount     =   2
               BeginProperty Column00 
                  DataField       =   "description"
                  Caption         =   "Hazard"
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
                  DataField       =   "rating"
                  Caption         =   "Rating"
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
               SplitCount      =   1
               BeginProperty Split0 
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                     Alignment       =   1
                  EndProperty
               EndProperty
            End
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00EAAF6F&
            Height          =   2475
            Left            =   -74925
            ScaleHeight     =   2415
            ScaleWidth      =   11055
            TabIndex        =   51
            Top             =   365
            Width           =   11115
            Begin VB.TextBox TxtReceipt 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   5
               Left            =   4215
               Locked          =   -1  'True
               TabIndex        =   44
               Top             =   885
               Width           =   2265
            End
            Begin VB.TextBox TxtReceipt 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   4
               Left            =   4215
               Locked          =   -1  'True
               TabIndex        =   43
               Top             =   570
               Width           =   2265
            End
            Begin VB.TextBox TxtReceipt 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   3
               Left            =   4215
               Locked          =   -1  'True
               TabIndex        =   42
               Top             =   255
               Width           =   2265
            End
            Begin VB.TextBox TxtReceipt 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   2
               Left            =   1860
               Locked          =   -1  'True
               TabIndex        =   41
               Top             =   885
               Width           =   2265
            End
            Begin VB.TextBox TxtReceipt 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   1
               Left            =   1860
               Locked          =   -1  'True
               TabIndex        =   40
               Top             =   570
               Width           =   2265
            End
            Begin VB.TextBox TxtReceipt 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   0
               Left            =   1860
               Locked          =   -1  'True
               TabIndex        =   39
               Top             =   255
               Width           =   2265
            End
            Begin VB.TextBox TxtAveragePrice 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   300
               Left            =   1860
               Locked          =   -1  'True
               TabIndex        =   47
               Top             =   1950
               Width           =   2265
            End
            Begin VB.TextBox TxtLastPrice 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   300
               Left            =   1860
               Locked          =   -1  'True
               TabIndex        =   46
               Top             =   1635
               Width           =   2265
            End
            Begin VB.TextBox TxtActualPrice 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   300
               Left            =   1860
               Locked          =   -1  'True
               TabIndex        =   45
               Top             =   1320
               Width           =   2265
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qty Received"
               ForeColor       =   &H80000005&
               Height          =   195
               Index           =   14
               Left            =   4680
               TabIndex        =   59
               Top             =   15
               Width           =   975
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qty Issued"
               ForeColor       =   &H80000005&
               Height          =   195
               Index           =   13
               Left            =   2250
               TabIndex        =   58
               Top             =   15
               Width           =   795
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Last 90 Days"
               ForeColor       =   &H80000005&
               Height          =   195
               Index           =   12
               Left            =   210
               TabIndex        =   57
               Top             =   915
               Width           =   930
            End
            Begin VB.Line Line1 
               Index           =   18
               X1              =   210
               X2              =   1875
               Y1              =   1170
               Y2              =   1170
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Last 60 Days"
               ForeColor       =   &H80000005&
               Height          =   195
               Index           =   11
               Left            =   210
               TabIndex        =   56
               Top             =   600
               Width           =   930
            End
            Begin VB.Line Line1 
               Index           =   17
               X1              =   210
               X2              =   1875
               Y1              =   855
               Y2              =   855
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Last 30 Days"
               ForeColor       =   &H80000005&
               Height          =   195
               Index           =   10
               Left            =   210
               TabIndex        =   55
               Top             =   285
               Width           =   930
            End
            Begin VB.Line Line1 
               Index           =   16
               X1              =   210
               X2              =   1875
               Y1              =   540
               Y2              =   540
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Average Price"
               ForeColor       =   &H80000005&
               Height          =   195
               Index           =   27
               Left            =   240
               TabIndex        =   54
               Top             =   1980
               Width           =   1005
            End
            Begin VB.Line Line1 
               Index           =   25
               X1              =   240
               X2              =   1905
               Y1              =   2235
               Y2              =   2235
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Last Price"
               ForeColor       =   &H80000005&
               Height          =   195
               Index           =   24
               Left            =   240
               TabIndex        =   53
               Top             =   1665
               Width           =   690
            End
            Begin VB.Line Line1 
               Index           =   22
               X1              =   240
               X2              =   1905
               Y1              =   1920
               Y2              =   1920
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Actual Price"
               ForeColor       =   &H80000005&
               Height          =   195
               Index           =   1
               Left            =   240
               TabIndex        =   52
               Top             =   1350
               Width           =   840
            End
            Begin VB.Line Line1 
               Index           =   1
               X1              =   240
               X2              =   1905
               Y1              =   1605
               Y2              =   1605
            End
         End
         Begin VB.PictureBox Picture4 
            Height          =   2475
            Left            =   -74925
            ScaleHeight     =   2415
            ScaleWidth      =   11055
            TabIndex        =   50
            Top             =   360
            Width           =   11115
            Begin MSDataGridLib.DataGrid Dgdetail 
               Height          =   2370
               Left            =   0
               TabIndex        =   38
               Top             =   0
               Width           =   11055
               _ExtentX        =   19500
               _ExtentY        =   4180
               _Version        =   393216
               AllowUpdate     =   0   'False
               BackColor       =   16777215
               BorderStyle     =   0
               HeadLines       =   1
               RowHeight       =   15
               RowDividerStyle =   6
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
               ColumnCount     =   5
               BeginProperty Column00 
                  DataField       =   "Kode Barang"
                  Caption         =   "Kode Barang"
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
                  DataField       =   "Nama Barang"
                  Caption         =   "Nama Barang"
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
                  DataField       =   "Jenis"
                  Caption         =   "Jenis"
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
                  DataField       =   "Kd. Brg Supplier"
                  Caption         =   "Kd. Brg Supplier"
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
                  DataField       =   "Unit"
                  Caption         =   "Unit"
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
                  MarqueeStyle    =   4
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
         End
         Begin VB.PictureBox Picture3 
            Height          =   2475
            Left            =   75
            ScaleHeight     =   2415
            ScaleWidth      =   11055
            TabIndex        =   49
            Top             =   360
            Width           =   11115
            Begin MSDataGridLib.DataGrid GridItem 
               Height          =   2415
               Left            =   0
               TabIndex        =   37
               Tag             =   "Partner"
               Top             =   0
               Width           =   11055
               _ExtentX        =   19500
               _ExtentY        =   4260
               _Version        =   393216
               AllowUpdate     =   0   'False
               BackColor       =   16777215
               BorderStyle     =   0
               HeadLines       =   1
               RowHeight       =   15
               RowDividerStyle =   6
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
               ColumnCount     =   10
               BeginProperty Column00 
                  DataField       =   "NoItem"
                  Caption         =   "Kode"
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
                  DataField       =   "WareHouse"
                  Caption         =   "Warehouse"
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
                  DataField       =   "NoGroup"
                  Caption         =   "Group"
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
                  DataField       =   "internalName"
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
               BeginProperty Column04 
                  DataField       =   "Merk"
                  Caption         =   "Kategori"
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
                  DataField       =   "Serial Supplier"
                  Caption         =   "Supplier PartNo"
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
                  DataField       =   "UOM"
                  Caption         =   "Satuan"
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
                  DataField       =   "MinStock"
                  Caption         =   "Min Stock"
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
               BeginProperty Column08 
                  DataField       =   "MaxStock"
                  Caption         =   "Max Stock"
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
               BeginProperty Column09 
                  DataField       =   "LeadTimeDays"
                  Caption         =   "Lead Time"
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
                     Object.Visible         =   0   'False
                  EndProperty
                  BeginProperty Column02 
                     Object.Visible         =   0   'False
                  EndProperty
                  BeginProperty Column03 
                  EndProperty
                  BeginProperty Column04 
                  EndProperty
                  BeginProperty Column05 
                     Object.Visible         =   0   'False
                  EndProperty
                  BeginProperty Column06 
                     ColumnWidth     =   1080
                  EndProperty
                  BeginProperty Column07 
                     Alignment       =   1
                  EndProperty
                  BeginProperty Column08 
                     Alignment       =   1
                     ColumnWidth     =   1080
                  EndProperty
                  BeginProperty Column09 
                     Alignment       =   1
                  EndProperty
               EndProperty
            End
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Currency"
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
         Index           =   23
         Left            =   5340
         TabIndex        =   81
         Top             =   2280
         Width           =   720
      End
      Begin VB.Line Line1 
         Index           =   21
         X1              =   5295
         X2              =   6960
         Y1              =   2535
         Y2              =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Based UOM"
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
         Index           =   22
         Left            =   5340
         TabIndex        =   80
         Top             =   1200
         Width           =   930
      End
      Begin VB.Line Line1 
         Index           =   20
         X1              =   5295
         X2              =   6960
         Y1              =   1461
         Y2              =   1461
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Konversi"
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
         Index           =   21
         Left            =   5340
         TabIndex        =   79
         Top             =   840
         Width           =   675
      End
      Begin VB.Line Line1 
         Index           =   19
         X1              =   5295
         X2              =   6960
         Y1              =   1095
         Y2              =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kategori"
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
         Index           =   15
         Left            =   150
         TabIndex        =   78
         Top             =   1215
         Width           =   675
      End
      Begin VB.Line Line1 
         Index           =   15
         X1              =   5295
         X2              =   6960
         Y1              =   4335
         Y2              =   4335
      End
      Begin VB.Line Line1 
         Index           =   14
         X1              =   105
         X2              =   1785
         Y1              =   3930
         Y2              =   3930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory Type"
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
         Index           =   9
         Left            =   150
         TabIndex        =   77
         Top             =   3690
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Vendor"
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
         Index           =   8
         Left            =   150
         TabIndex        =   76
         Top             =   1935
         Width           =   1110
      End
      Begin VB.Line Line1 
         Index           =   13
         X1              =   105
         X2              =   1770
         Y1              =   2190
         Y2              =   2190
      End
      Begin VB.Line Line1 
         Index           =   12
         X1              =   5295
         X2              =   6960
         Y1              =   2175
         Y2              =   2175
      End
      Begin VB.Line Line1 
         Index           =   11
         X1              =   135
         X2              =   1800
         Y1              =   2910
         Y2              =   2910
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   105
         X2              =   1770
         Y1              =   1830
         Y2              =   1830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lokasi Gudang"
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
         Index           =   7
         Left            =   150
         TabIndex        =   75
         Top             =   1590
         Width           =   1170
      End
      Begin VB.Line Line1 
         Index           =   9
         X1              =   5295
         X2              =   6960
         Y1              =   3975
         Y2              =   3975
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   5295
         X2              =   6960
         Y1              =   3615
         Y2              =   3615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reorder Level"
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
         Index           =   4
         Left            =   5340
         TabIndex        =   74
         Top             =   3720
         Width           =   1125
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   5295
         X2              =   6960
         Y1              =   3255
         Y2              =   3255
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   135
         X2              =   1800
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   5295
         X2              =   6960
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   5295
         X2              =   6960
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   135
         X2              =   1800
         Y1              =   2572
         Y2              =   2572
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   135
         X2              =   1800
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Line ln 
         Index           =   0
         X1              =   135
         X2              =   1800
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   135
         X2              =   1800
         Y1              =   390
         Y2              =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Part Number"
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
         Left            =   150
         TabIndex        =   73
         Top             =   150
         Width           =   1035
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Inventory"
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
         Left            =   150
         TabIndex        =   72
         Top             =   510
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kelompok"
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
         Left            =   150
         TabIndex        =   71
         Top             =   870
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor ItemID"
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
         Index           =   3
         Left            =   150
         TabIndex        =   70
         Top             =   2295
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase UOM"
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
         Index           =   5
         Left            =   5340
         TabIndex        =   69
         Top             =   480
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min. Stock"
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
         Index           =   6
         Left            =   5340
         TabIndex        =   68
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max. Stock"
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
         Index           =   16
         Left            =   5340
         TabIndex        =   67
         Top             =   3000
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga/Unit"
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
         Index           =   17
         Left            =   5340
         TabIndex        =   66
         Top             =   1920
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reorder Qty"
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
         Index           =   18
         Left            =   5340
         TabIndex        =   65
         Top             =   3360
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lead Time"
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
         Index           =   19
         Left            =   5340
         TabIndex        =   64
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Index           =   20
         Left            =   150
         TabIndex        =   63
         Top             =   2670
         Width           =   525
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Internal"
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
         Left            =   5340
         TabIndex        =   62
         Top             =   120
         Width           =   1140
      End
      Begin VB.Line ln 
         Index           =   1
         X1              =   5295
         X2              =   6960
         Y1              =   375
         Y2              =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales UOM"
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
         Index           =   25
         Left            =   5340
         TabIndex        =   61
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serial / Lot Tracking"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   26
         Left            =   150
         TabIndex        =   60
         Top             =   3030
         Width           =   1410
      End
      Begin VB.Line Line1 
         Index           =   23
         X1              =   135
         X2              =   1800
         Y1              =   3270
         Y2              =   3270
      End
      Begin VB.Line Line1 
         Index           =   24
         X1              =   5340
         X2              =   7005
         Y1              =   1815
         Y2              =   1815
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   8280
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmItemData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RcGudang As New DBQuick
Private RcGroup As New DBQuick
Private RcBOM As New DBQuick
Private MyData As New clsTransaksi
Private mAdd As Boolean
Private mLead As Boolean
Private mKeyLoad As Boolean
Private RcPartner As New DBQuick
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private mLastGDG As String
Private mVarOldLead As Long
Private RsHazard As New DBQuick
Private RsLoadHazard As New DBQuick
Private IDGen As New IDGenerator
Private lMode As String

Public Property Let SetMode(Value As String)
   lMode = Value
End Property

Private Sub SimpanHazard()
On Error GoTo xErr
With RsHazard.DBRecordset
   If .Recordcount > 0 Then
      If SendDataToServer("delete from item_hazard_line where noItem ='" & MyDDE.GetFieldByName("noItem") & "'") = True Then
         .MoveFirst
         While Not .EOF
            SendDataToServer "insert into item_hazard_line(NoItem,code_hazard,rating) values ('" & _
                             MyDDE.GetFieldByName("NoItem") & "','" & .Fields("code_hazard") & "'," & .Fields("Rating") & ")"
            .MoveNext
         Wend
      End If
   End If
End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub Check1_Click()
   If Check1.Value = 1 Then
      Check2.Value = 0
      Check3.Value = 0
   End If
End Sub

Private Sub Check2_Click()
   If Check2.Value = 1 Then
      Check1.Value = 0
      Check3.Value = 0
   End If
End Sub

Private Sub Check3_Click()
   If Check3.Value = 1 Then
      Check2.Value = 0
      Check1.Value = 0
   End If
End Sub

Private Sub cmdLink_Click(Index As Integer)
   OpenPartner Index
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Command1_Click(Index As Integer)
If MyDDE.ActiveRecordset Is Nothing Then Exit Sub
Select Case Index
       Case 0:
            If MyDDE.ActiveRecordset.State = 1 Then
            If MyDDE.ActiveRecordset.Recordcount <> 0 Then
               'If MyDDE.ActiveRecordset.Fields("Manufacture") = False Then
                  FrmAvaillable.SetNoItem(txtBox(1)) = txtBox(0)
                  FrmAvaillable.Show vbModal
               'End If
            End If
            End If
       Case 1:
            FrmItemReference.SetFocus
            If FrmItemReference.MyDDE.ActiveRecordset.Recordcount <> 0 Then
               FrmItemReference.MyDDE.FindStringData "[Kode Barang] ='" & txtBox(0) & "'"
               FrmItemReference.SSTab1.Tab = 1
            End If
End Select
End Sub

Private Sub GridItem_Error(ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub

Private Sub GridItem_KeyDown(KeyCode As Integer, Shift As Integer)
Form_KeyDown KeyCode, Shift
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If mKeyLoad = False Then mKeyLoad = True Else mKeyLoad = False
'If mKeyLoad = False Then
If mAdd = False Then ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()

Select Case UCase(lMode)
   Case "", "CARD":
      MyDDE.SetPermissions = aksess.MayDo("Kartu Stok")
      
      '* Nama Inventory *'
      lbl(0).Visible = False
      ln(0).Visible = False
      txtBox(1).Visible = False
      
      '* Nama Internal *'
      lbl(1).Visible = True
      ln(1).Visible = True
      txtBox(14).Visible = True
      
   Case "MANAGER":
      MyDDE.SetPermissions = aksess.MayDo("Kartu Stok Manager")
      
      '* Nama Inventory *'
      lbl(0).Visible = True
      ln(0).Visible = True
      txtBox(1).Visible = True
      
      '* Nama Internal *'
      lbl(1).Visible = True
      ln(1).Visible = True
      txtBox(14).Visible = True
      
   Case "PURCHASING":
      MyDDE.SetPermissions = aksess.MayDo("Kartu Stok Purchasing")
      
      '* Nama Inventory *'
      lbl(0).Visible = True
      ln(0).Visible = True
      txtBox(1).Visible = True
      
      '* Nama Internal *'
      lbl(1).Visible = False
      ln(1).Visible = False
      txtBox(14).Visible = False
End Select

SSTab1.Tab = 0
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
GridLayout
Set mCall = New frmCaller
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmItemData
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
    .PrepareQuery = " SELECT Inventory.NoItem, Inventory.ItemName, Inventory.Merk, Inventory.[Serial Supplier], Inventory.UOM, Inventory.UOMKonversi, Inventory.UOMPurchase, " & _
                    " Inventory.MinStock, Inventory.MaxStock, Inventory.StatusItem, Inventory.ROP, Inventory.ROL, Inventory.WareHouse AS [Kode Gudang Persediaan]," & _
                    " WareHouse.[WareHouse Name] AS [Nama Gudang Persediaan], Inventory.NoAccount AS [Kode Perkiraan], Inventory.NoGroup AS [Kode Kelompok]," & _
                    " [Inventory Group].[Group Name] AS [Kelompok Persediaan], Inventory.Status, Inventory.PriceIn AS Harga, Inventory.LeadTimeDays," & _
                    " Inventory.PartnerID AS PartnerID, PartnerDB.CompanyName, Inventory.Manufacture, [UOM Table].Description AS [UOM Description],Inventory.CurrID, " & _
                    " InternalName,Tracking_code,UOMSales,item_tracking_code.description as TrackingName, Inventory.CategID, inventory_categories.description AS category " & _
                    " FROM Inventory INNER JOIN [Inventory Group] ON Inventory.NoGroup = [Inventory Group].NoGroup " & _
                    " INNER JOIN WareHouse ON Inventory.WareHouse = WareHouse.WareHouse  " & _
                    " INNER JOIN [UOM Table] ON Inventory.UOM = [UOM Table].UOM  " & _
                    " LEFT OUTER JOIN inventory_categories ON Inventory.CategID = inventory_categories.categid AND " & _
                    " Inventory.nogroup = inventory_categories.nogroup And [Inventory Group].nogroup = inventory_categories.nogroup " & _
                    " LEFT OUTER JOIN PartnerDB ON Inventory.PartnerID = PartnerDB.PartnerID  " & _
                    " left outer join item_tracking_code on item_tracking_code.code = inventory.tracking_code"
                    
'    Debug.Print .PrepareQuery
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MyDDE.CheckRecordPendinged = True Then
   ScanKey vbKeyF5, 0, MyDDE
   If MyDDE.IsSucces = True Then
      Cancel = False
      MyDDE.ClearRecordset
      RcGudang.CloseDB
      RcGroup.CloseDB
   Else
      Cancel = True
   End If
Else
   
   MyDDE.ClearRecordset
      RcGudang.CloseDB
      RcGroup.CloseDB
End If
End Sub

Private Sub Form_Resize()


Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmItemData = Nothing
End Sub

Private Sub mCall_BeforeUnload()
    Select Case UCase(mCall.Tag)
        Case "KELOMPOK PERSEDIAAN":
             If cmdLink(1).Enabled = True Then cmdLink(1).SetFocus
             IDGen.ExtParameter = MyDDE.GetFieldByName("Kode Kelompok") & "-" & MyDDE.GetFieldByName("CategID")
             txtBox(0).Text = IDGen.GetID("IN")
             
            If MyDDE.GetFieldByName("kode kelompok") = "FG" Then
              'CmdLink(2).Enabled = False
              'txtBox(3).Enabled = False
              Dim rsCek As New DBQuick
              rsCek.DBOpen "select partnerID from partnerdb where [default]=1", CNN
              If rsCek.DBRecordset.Recordcount > 0 Then
               MyDDE.GetFieldByName("PartnerID") = rsCek.DBRecordset.Fields(0)
               MyDDE.GetFieldByName("serial_supplier") = rsCek.DBRecordset.Fields(0)
               MyDDE.GetFieldByName("CompanyName") = "-"
              End If
              rsCek.CloseDB
            Else
              cmdLink(2).Enabled = True
              txtBox(3).Enabled = True
            End If

        Case "KELOMPOK GUDANG PERSEDIAAN":
             If txtBox(3).Enabled = True Then cmdLink(2).SetFocus
        Case "MASTER PARTNER":
             If txtBox(3).Enabled = True Then txtBox(3).SetFocus
        Case "MASTER UOM":
             If txtBox(12).Enabled = True Then txtBox(12).SetFocus
        Case "PURCHASE UOM":
             If txtBox(4).Enabled = True Then txtBox(4).SetFocus
        Case "MASTER CURRENCY":
             If txtBox(13).Enabled = True Then txtBox(13).SetFocus
        Case "KATEGORI BARANG - " & UCase(MyDDE.GetFieldByName("Kelompok Persediaan")):
             If cmdLink(8).Enabled = True Then txtBox(2).SetFocus
             IDGen.ExtParameter = MyDDE.GetFieldByName("Kode Kelompok") & "-" & MyDDE.GetFieldByName("CategID")
             txtBox(0).Text = IDGen.GetID("IN")
             MyDDE.GetFieldByName("noItem") = txtBox(0).Text
    End Select
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
With MyDDE
     Select Case UCase(TagForm)
            Case "KELOMPOK PERSEDIAAN":
                .GetFieldByName("Kode Kelompok") = mCall.GetFieldByName(0)
                .GetFieldByName("Kelompok Persediaan") = mCall.GetFieldByName(1)
                mLastGDG = mCall.GetFieldByName(0)
                'If mAdd = True Then MyDDE.GetFieldByName("NoItem") = MyData.PrepareIndex(tmbInventory, 10 - Len(mCall.GetFieldByName(0)), mCall.GetFieldByName(0), mCall.GetFieldByName(0) & "-")
            Case "KELOMPOK GUDANG PERSEDIAAN":
                 .GetFieldByName("Kode Gudang Persediaan") = mCall.GetFieldByName(0)
                 .GetFieldByName("Nama Gudang Persediaan") = mCall.GetFieldByName(1)
                 .GetFieldByName("Kode Perkiraan") = mCall.GetFieldByName(2)
            Case "MASTER PARTNER":
                 .GetFieldByName("PartnerID") = mCall.GetFieldByName(0)
                 .GetFieldByName("CompanyName") = mCall.GetFieldByName(1)
            Case "MASTER UOM":
                 .GetFieldByName("UOM") = mCall.GetFieldByName(0)
            Case "PURCHASE UOM":
                 .GetFieldByName("UOMPurchase") = mCall.GetFieldByName(0)
                 .GetFieldByName("UOM") = mCall.GetFieldByName(0)
                 .GetFieldByName("UOMSales") = mCall.GetFieldByName(0)
            Case "MASTER CURRENCY":
                 .GetFieldByName("Currid") = mCall.GetFieldByName(0)
            Case "HAZARD":
                  MyDDE.ChildRecordset.MoveLast
                  MyDDE.ChildRecordset.Fields("code_hazard") = mCall.GetFieldByName("code")
                  MyDDE.ChildRecordset.Fields("description") = mCall.GetFieldByName("description")
                  MyDDE.ChildRecordset.Fields("Rating") = 0
            Case "TRACKING CODE":
                  MyDDE.GetFieldByName("tracking_code") = mCall.GetFieldByName("code")
                  MyDDE.GetFieldByName("TrackingName") = mCall.GetFieldByName(1)
            Case "SALES UOM": MyDDE.GetFieldByName("UOMSales") = mCall.GetFieldByName(0)
            Case "KATEGORI BARANG - " & UCase(MyDDE.GetFieldByName("Kelompok Persediaan")):
                .GetFieldByName("CategID") = mCall.GetFieldByName(0)
                .GetFieldByName("category") = mCall.GetFieldByName(1)
                mLastGDG = mCall.GetFieldByName(0)
'                If mAdd = True Then
'                    MyDDE.GetFieldByName("NoItem") = MyData.PrepareIndex(tmbInventory, _
'                    10 - Len(mCall.GetFieldByName(0)), _
'                    mCall.GetFieldByName(0), mCall.GetFieldByName(0) & "-")
'                End If
        
     End Select
End With
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Dim strSQL As String
Dim x As Integer

Select Case AdReasonActiveDb
       Case tmbAddNew:
         Check1.Enabled = True
         Check2.Enabled = True
         Check3.Enabled = True
            
            SetelKosong
            txtBox(0).Enabled = False
            
            IDGen.ExtParameter = Label2(0).Text & "-" & txtBox(2).Text
            
            With MyDDE
               .GetFieldByName("NoItem") = IDGen.GetID("IN")  'MyData.PrepareIndex(tmbInventory, 10 - Len(mCall.GetFieldByName(0)), mCall.GetFieldByName(0), mCall.GetFieldByName(0) & "-")
               .GetFieldByName("uom") = "PCS"
               .GetFieldByName("uomPurchase") = "PCS"
               .GetFieldByName("uomSales") = "PCS"
               .GetFieldByName("rop") = 0
               .GetFieldByName("rol") = 0
               .GetFieldByName("LeadTimeDays") = 0
               .GetFieldByName("UOMKonversi") = 1
               .GetFieldByName("Harga") = 0
               .GetFieldByName("CurrID") = "IDR"
               .GetFieldByName("MinStock") = 1
               .GetFieldByName("maxStock") = 1000
               .GetFieldByName("ROP") = 100
               .GetFieldByName("ROL") = 1
            End With
            
            Check2.Value = 1
            mVarOldLead = 0
            mAdd = True
            Combo1.ListIndex = 0
       Case tmbEdit:
        
         Check1.Enabled = True
         Check2.Enabled = True
         Check3.Enabled = True
            txtBox(0).Enabled = False
'            txtBox(1).SetFocus
'            cboGudang(1).Enabled = False
            
'            txtBox(0).Enabled = False
 '           If txtBox(1).Enabled = True Then txtBox(1).SetFocus
            mVarOldLead = CDbl(txtBox(10))
            mAdd = True
       Case tmbSave:
            mAdd = False
            If MyDDE.IsChildMemberReady = True Then SimpanHazard
            
       Case tmbPrint:
         Select Case UCase(lMode)
            Case "", "CARD":
               CallRPTReport "Tabel Item.rpt"
            Case "MANAGER":
               Dim cRep As New utility
               cRep.CallReportView "select * from [tabel item]", "Tabel Item Manager.rpt", ReportPos
            Case "PURCHASING":
               cRep.CallReportView "select * from [tabel item]", "Tabel Item Purc.rpt", ReportPos
         End Select
       Case tmbDetail: OpenPartner 9
       Case Else: 'mVarDataDc = False
End Select


'messagebox MyData.PrepareIndex(tmbInventory, 15 - Len(cboGudang(1).BoundText), cboGudang(1).BoundText, cboGudang(1).BoundText & "/")
mAdd = txtBox(1).Enabled

For x = 0 To 8
   cmdLink(x).Enabled = txtBox(1).Enabled
Next

txtBox(4).Enabled = False
txtBox(12).Enabled = False
txtBox(13).Enabled = False
mLead = mAdd
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
'On Error Resume Next
   Select Case AdReasonActiveDb
      Case tmbSave, tmbDelete: PrepareQuery
   End Select
'Err.Clear
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If MyDDE.ActiveRecordset.Recordcount <> 0 Then
   If MyDDE.ActiveRecordset.Fields("Manufacture") = False Then
      Command1(1).Enabled = True
   Else
      Command1(1).Enabled = False
   End If
   Check1.Value = 0
   Check2.Value = 0
   Check3.Value = 0
   Select Case MyDDE.GetFieldByName("Manufacture")
      Case 0: Check2.Value = 1
      Case 1: Check1.Value = 1
      Case 2: Check3.Value = 1
   End Select
Else
   Command1(1).Enabled = False
End If

   RsHazard.DBOpen "SELECT item_hazard_line.code_hazard, tbl_hazard.description, item_hazard_line.rating " & _
                   " FROM item_hazard_line INNER JOIN tbl_hazard ON item_hazard_line.code_hazard = tbl_hazard.code " & _
                   " where item_hazard_line.NoItem ='" & MyDDE.GetFieldByName("NoItem") & "'", CNN

Call SSTab1_Click(SSTab1.Tab)


End Sub

'
Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim mDel As New clsDelete
Dim CascaDel As String
On Error GoTo Hell
Select Case AdReasonActiveDb
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               If mDel.CekDelete(txtBox(0), reDelMasterBarang) = False Then
                  MyDDE.IsChildMemberReady = True
                  PrepareQuery
               Else
                  MyDDE.CancelTrans = True
                  CascaDel = mDel.CekCascadeDeleteTable(txtBox(0), reDelMasterBarang)
                  MessageBox "Data ' " & txtBox(1).Text & " (" & txtBox(0) & ") '  sedang digunakan pada transaksi " & _
                  UCase(CascaDel) & vbCrLf & "Record tidak bisa dihapus..", "Kontrol Penghapusan Data", msgOkOnly, msgExclamation
                  MyDDE.IsChildMemberReady = False
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
       
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If MyDDE.GetFieldByName("MinStock") >= MyDDE.GetFieldByName("MaxStock") Then
                  MyDDE.IsChildMemberReady = False
                  MessageBox "Minimum stok tidak boleh lebih besar dari maximum stok.", "Stok Level", msgOkOnly, msgCrtical
               Else
                  MyDDE.IsChildMemberReady = True
               End If
               
'               MyDDE.GetFieldByName("WareHouse") = CboGudang(0).BoundText
'               MyDDE.GetFieldByName("NoGroup") = CboGudang(1).BoundText
               'PrepareQuery
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
Set mDel = Nothing
Exit Sub
Hell:
   MessageBox Err.Number, "Error", msgOkOnly, msgExclamation
   Err.Clear

End Sub



Private Sub SSTab1_Click(PreviousTab As Integer)
MyDDE.InitControlSet = [Master Mode]
Select Case SSTab1.Tab
       Case 0: RcBOM.CloseDB
       Case 1: WhereUsedBom
       Case 2:
            OpenQTY
            HitungAVG txtBox(0)
            RcBOM.CloseDB
       Case 3: MyDDE.InitControlSet = [Transaction Mode]
               Set MyDDE.ChildRecordset = RsHazard.DBRecordset
               Set GridHazard.DataSource = MyDDE.ChildRecordset
End Select
End Sub

Private Sub txtBox_Change(Index As Integer)
Dim I As Integer
If Index = 10 And mLead = True Then
   I = MessageBox("Anda akan merubah Lead Time yang telah dikalkulasi oleh system", "Peringatan", msgYesNo)
   mLead = False
   If I = 1 Then
      
   Else
      txtBox(10) = mVarOldLead
   End If
End If
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub PrepareQuery()
Dim stManufacture As Integer
On Error GoTo xErr
If Check1.Value = 1 Then
   stManufacture = 1
ElseIf Check2.Value = 1 Then
   stManufacture = 0
ElseIf Check3.Value = 1 Then
   stManufacture = 2
End If

With MyDDE
    .PrepareAppend = " INSERT INTO Inventory ( PartnerID,NoItem,PriceIn , NoGroup,  ItemName, Merk, [Serial Supplier], UOM, MinStock, MaxStock, StatusItem,rol,rop,NoAccount,Warehouse,Status,LeadTimeDays,UOMPurchase,UOMKonversi,CurrID,Manufacture,InternalName,tracking_code,UOMSales, CategID) " & _
                     " VALUES (N'" & IIf(IsNull(MyDDE.GetFieldByName("PartnerID")), "", MyDDE.GetFieldByName("PartnerID")) & "',N'" & txtBox(0) & "'," & CDbl(txtBox(9)) & ",  N'" & IIf(IsNull(MyDDE.GetFieldByName("Kode Kelompok")), "", MyDDE.GetFieldByName("Kode Kelompok")) & "', N'" & txtBox(1) & "', N'" & ValidString(txtBox(2)) & "', N'" & ValidString(txtBox(3)) & "', N'" & ValidString(txtBox(4)) & "' ," & _
                     "  " & CDbl(txtBox(5)) & ", " & CDbl(txtBox(6)) & ",  N'ITEM'," & CDbl(txtBox(7)) & "," & CDbl(txtBox(8)) & ",N'" & IIf(IsNull(MyDDE.GetFieldByName("Kode Perkiraan")), "", MyDDE.GetFieldByName("Kode Perkiraan")) & "',N'" & IIf(IsNull(MyDDE.GetFieldByName("Kode Gudang Persediaan")), "", MyDDE.GetFieldByName("Kode Gudang Persediaan")) & "',N'" & Combo1.Text & "'," & CDbl(txtBox(10)) & "," & _
                     " N'" & ValidString(txtBox(4)) & "'," & CDbl(txtBox(11)) & ",N'" & txtBox(13) & "'," & stManufacture & ",'" & IIf(IsNull(.GetFieldByName("InternalName")), "", .GetFieldByName("InternalName")) & "','" & IIf(IsNull(.GetFieldByName("Tracking_code")), "", .GetFieldByName("Tracking_code")) & "','" & IIf(IsNull(.GetFieldByName("UOMsales")), "", .GetFieldByName("UOMSales")) & "','" & IIf(IsNull(.GetFieldByName("CategID")), "", .GetFieldByName("CategID")) & "')"
                     

    
    .PrepareUpdate = " UPDATE Inventory Set Currid=N'" & txtBox(13) & "',UOMPurchase =N'" & ValidString(txtBox(4)) & _
                       "' ,UOMKonversi = " & CDbl(txtBox(11)) & ",LeadTimeDays =  " & CDbl(txtBox(10)) & _
                       " ,PartnerID =N'" & MyDDE.GetFieldByName("PartnerID") & "', PriceIn=" & CDbl(txtBox(9)) & _
                       ",Status =N'" & Combo1.Text & "',NoGroup = N'" & MyDDE.GetFieldByName("Kode Kelompok") & _
                       "', ItemName = N'" & txtBox(1) & "', Merk = N'" & ValidString(txtBox(2)) & _
                       "', [Serial Supplier] = N'" & ValidString(txtBox(3)) & "', UOM = N'" & _
                       ValidString(txtBox(12)) & "', MinStock = " & CDbl(txtBox(5)) & ",rol =" & _
                       CDbl(txtBox(7)) & ",rop = " & CDbl(txtBox(8)) & ",NoAccount='" & _
                       MyDDE.GetFieldByName("Kode Perkiraan") & "',Warehouse='" & _
                       MyDDE.GetFieldByName("Kode Gudang Persediaan") & "', MaxStock = " & CDbl(txtBox(6)) & _
                       ",Manufacture = " & stManufacture & ",InternalName='" & .GetFieldByName("InternalName") & _
                       "', Tracking_code ='" & .GetFieldByName("Tracking_code") & "', UOMSales ='" & .GetFieldByName("UOMSales") & "', CategID ='" & .GetFieldByName("CategID") & "'" & _
                     " WHERE (NoItem = N'" & ValidString(txtBox(0)) & "')"
'    Debug.Print .PrepareUpdate
    .PrepareDelete = " DELETE FROM Inventory WHERE    (NoItem = N'" & ValidString(txtBox(0)) & "')"
End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub txtBox_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
       Case 5, 6, 7, 8, 11:
            ValidNum KeyAscii
       Case Else:
End Select
End Sub

Private Sub SetelKosong()
With MyDDE
     .GetFieldByName("ItemName") = "-"
     .GetFieldByName("Merk") = "-"
     .GetFieldByName("Serial Supplier") = "-"
     .GetFieldByName("UOM") = "KG"
     .GetFieldByName("MinStock") = 0
     .GetFieldByName("MaxStock") = 0
     .GetFieldByName("Manufacture") = 1
     .GetFieldByName("Purchase") = 0
     .GetFieldByName("SemiManufacture") = 0
End With
End Sub

Private Function OpenPartner(ByVal Index As Integer) As Boolean
Dim strSQL As String
Select Case Index
    Case 0: RcPartner.DBOpen " SELECT NoGroup AS [Kode Kelompok Persediaan], [Group Name] AS [Kelompok Persediaan] FROM         [Inventory Group]", CNN, lckLockReadOnly
    Case 1: RcPartner.DBOpen " SELECT WareHouse.WareHouse AS [Kode Gudang Persediaan], WareHouse.[WareHouse Name] AS [Nama Gudang Persediaan],                        WareHouse.NoAccount AS [Kode Perkiraan] FROM         [Inventory Group] INNER JOIN                       GLAccount ON [Inventory Group].NoAccount = GLAccount.NoAccount INNER JOIN                       WareHouse ON GLAccount.NoAccount = WareHouse.GroupAccount WHERE     ([Inventory Group].NoGroup = N'" & MyDDE.GetFieldByName("Kode Kelompok") & "')", CNN, lckLockReadOnly
    Case 2: RcPartner.DBOpen " SELECT PartnerID AS [Kode Partner], CompanyName AS [Nama Perusahaan], Address AS Alamat, City AS [Kota Telp], Phone FROM         PartnerDB WHERE     (PartnerType = N'SUPPLIER') ORDER BY PartnerID", CNN, lckLockBatch
    Case 3, 4, 6: RcPartner.DBOpen " SELECT * FROM [UOM Table] ORDER BY UOM ", CNN, lckLockReadOnly
    Case 5: RcPartner.DBOpen " SELECT  CurrID, [Currency Name] FROM [Currency Setup] ORDER BY CurrID", CNN, lckLockBatch
    Case 7: RcPartner.DBOpen " select code,description from item_tracking_code ", CNN
    Case 8: RcPartner.DBOpen " SELECT categid, description From inventory_categories WHERE (nogroup = N'" & MyDDE.GetFieldByName("Kode Kelompok") & "') ORDER BY description", CNN, lckLockBatch
    Case 9: RcPartner.DBOpen "select code,description from tbl_hazard " & _
            SQLLookupParameter(RsHazard.DBRecordset, "code", "code_hazard"), CNN

End Select
'MessageBox "SELECT WareHouse.WareHouse AS [Kode Gudang Persediaan], WareHouse.[WareHouse Name] AS [Nama Gudang Persediaan],                        WareHouse.NoAccount AS [Kode Perkiraan] FROM         [Inventory Group] INNER JOIN                       GLAccount ON [Inventory Group].NoAccount = GLAccount.NoAccount INNER JOIN                       WareHouse ON GLAccount.NoAccount = WareHouse.GroupAccount WHERE     ([Inventory Group].NoGroup = N'" & MyDDE.GetFieldByName("Kode Kelompok") & "')"
If RcPartner.Recordcount <> 0 Then
    Select Case Index
           Case 0: mCall.FromTagActive = "KELOMPOK PERSEDIAAN"
           Case 1: mCall.FromTagActive = "KELOMPOK GUDANG PERSEDIAAN"
           Case 2: mCall.FromTagActive = "Master Partner"
           Case 3: mCall.FromTagActive = "Purchase UOM"
           Case 4: mCall.FromTagActive = "Master UOM"
           Case 5: mCall.FromTagActive = "Master Currency"
           Case 6: mCall.FromTagActive = "Sales UOM"
           Case 7: mCall.FromTagActive = "Tracking Code"
           Case 8: mCall.FromTagActive = "Kategori Barang - " & MyDDE.GetFieldByName("Kelompok Persediaan")
           Case 9: mCall.FromTagActive = "HAZARD"
    End Select
    Set mCall.FormData = RcPartner.DBRecordset
    mCall.LookUp Me
    'mLastGDG = ""
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgCrtical
   OpenPartner = True
End If

End Function

Private Sub GridLayout()
With GridItem
    .Columns(0).width = 1200
    .Columns(1).width = 1500
    .Columns(2).width = 1500
    .Columns(3).width = 4670
    .Columns(4).width = 1200
    .Columns(5).width = 1400
    .Columns(6).width = 850
    .Columns(7).width = 850
    .Columns(8).width = 850
    .Columns(9).width = 850
End With
With dgDetail
    .Columns(0).width = 1709.858
    .Columns(1).width = 3179.906
    .Columns(2).width = 2085.166
    .Columns(3).width = 1739.906
    .Columns(4).width = 1769.953
End With
End Sub

Private Sub txtBox_LostFocus(Index As Integer)
   Select Case Index
      Case 1: If (Trim(txtBox(14).Text) = "") Or (txtBox(14).Text = "-") Then txtBox(14).Text = txtBox(1).Text
      Case 14: If (Trim(txtBox(1).Text) = "") Or (txtBox(1).Text = "-") Then txtBox(1).Text = txtBox(14).Text
   End Select
End Sub

Private Sub WhereUsedBom()
RcBOM.DBOpen "SELECT     [BOM Component Detail].NoItem AS [Kode Barang], Inventory.ItemName AS [Nama Barang], Inventory.[Serial Supplier] AS [Kd. Brg Supplier],  Inventory.UOM AS Unit, Inventory.Merk AS Jenis FROM         [BOM Component Detail] INNER JOIN     Inventory ON [BOM Component Detail].NoItem = Inventory.NoItem WHERE     (Inventory.Manufacture = 1) AND ([BOM Component Detail].Component = N'" & MyDDE.GetFieldByName("NoItem") & "') GROUP BY [BOM Component Detail].NoItem, Inventory.ItemName, Inventory.UOM, Inventory.Merk, Inventory.[Serial Supplier] ORDER BY [BOM Component Detail].NoItem", CNN, lckLockReadOnly
Set dgDetail.DataSource = RcBOM.DBRecordset
End Sub

Private Sub OpenQTY()
Dim RcQty As New DBQuick
RcQty.DBOpen "SELECT     SUM(QTYI) AS QTYI, SUM(QTYII) AS QTYII, SUM(QTYIII) AS QTYIII FROM         [History Issued] WHERE     (NoItem = N'" & MyDDE.GetFieldByName("NoItem") & "')", CNN, lckLockReadOnly
With RcQty.DBRecordset
     If .Recordcount <> 0 Then
        TxtReceipt(0) = FormatNumber(IIf(Not IsNull(.Fields(0)), .Fields(0), 0), 0)
        TxtReceipt(1) = FormatNumber(IIf(Not IsNull(.Fields(1)), .Fields(1), 0), 0)
        TxtReceipt(2) = FormatNumber(IIf(Not IsNull(.Fields(2)), .Fields(2), 0), 0)
     Else
        TxtReceipt(0) = 0
        TxtReceipt(1) = 0
        TxtReceipt(2) = 0
     End If
End With
RcQty.DBOpen "SELECT     SUM(QTYI) AS QTYI, SUM(QTYII) AS QTYII, SUM(QTYIII) AS QTYIII FROM         [History Receipt] WHERE     (NoItem = N'" & MyDDE.GetFieldByName("NoItem") & "')", CNN, lckLockReadOnly
With RcQty.DBRecordset
     If .Recordcount <> 0 Then
        TxtReceipt(3) = FormatNumber(IIf(Not IsNull(.Fields(0)), .Fields(0), 0), 0)
        TxtReceipt(4) = FormatNumber(IIf(Not IsNull(.Fields(1)), .Fields(1), 0), 0)
        TxtReceipt(5) = FormatNumber(IIf(Not IsNull(.Fields(2)), .Fields(2), 0), 0)
     Else
        TxtReceipt(3) = 0
        TxtReceipt(4) = 0
        TxtReceipt(5) = 0
     End If
End With
RcQty.CloseDB
Set RcQty = Nothing
End Sub

Private Sub HitungAVG(ByVal NoItem As String)
On Error Resume Next
Dim RcTotQTY As New DBQuick
Dim RcTotHPP As New DBQuick
Dim RcPrice As New DBQuick

Dim HasilAVG As Currency

'Digunakan ambil tabel invntory
RcPrice.DBOpen "select noitem,fixcost,lastcost from inventory where (noitem='" & NoItem & "')", CNN, lckLockReadOnly

With RcPrice.DBRecordset
     TxtActualPrice = .Fields("fixcost")
     TxtLastPrice = .Fields("lastcost")
End With

'Digunakan hitung total sum HPP
RcTotHPP.DBOpen "select noitem,sum(qty_in * HPP) as Total_HPP from [inventory tabel] where (noitem='" & NoItem & "') group by noitem", CNN, lckLockReadOnly
'Digunakan Hitung total sum QTY_In
RcTotQTY.DBOpen "select noitem,sum(qty_in) as Total_QTY from [inventory tabel]  where (noitem='" & NoItem & "') group by noitem", CNN, lckLockReadOnly
If (RcTotHPP.Recordcount > 0) And (RcTotQTY.Recordcount > 0) Then
    HasilAVG = (RcTotHPP.DBRecordset.Fields("total_hpp") / RcTotQTY.DBRecordset.Fields("total_QTY"))

    TxtAveragePrice = HasilAVG
Else
    TxtAveragePrice = 0
End If

RcPrice.CloseDB
RcTotHPP.CloseDB
RcTotQTY.CloseDB

Set RcPrice = Nothing
Set RcTotHPP = Nothing
Set RcTotQTY = Nothing
End Sub


