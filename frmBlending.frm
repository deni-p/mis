VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmBlending 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blending Instruction"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBlending.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   11880
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5880
      Left            =   0
      ScaleHeight     =   5880
      ScaleWidth      =   11880
      TabIndex        =   33
      Top             =   0
      Width           =   11880
      Begin TabDlg.SSTab SSTab1 
         Height          =   5700
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   10054
         _Version        =   393216
         TabHeight       =   520
         BackColor       =   8438015
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Blending"
         TabPicture(0)   =   "frmBlending.frx":6852
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label1(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label1(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label1(3)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Line1(0)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Line1(1)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Line1(2)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Line1(3)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Line1(16)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label1(20)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Line1(17)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label1(19)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Line1(18)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label1(21)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Label4"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "txtDateBlending(1)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "txtDateBlending(0)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Frame1"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Frame2"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "gridBlending"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "cmdLink(0)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "cmdLink(1)"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "cmdLink(2)"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "txtBlending(0)"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "txtBlending(1)"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "txtBlending(4)"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "txtBlending(2)"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).ControlCount=   27
         TabCaption(1)   =   "Screw"
         TabPicture(1)   =   "frmBlending.frx":686E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "viewDate"
         Tab(1).Control(1)=   "gridScrew"
         Tab(1).Control(2)=   "txtScrew(1)"
         Tab(1).Control(3)=   "txtScrew(0)"
         Tab(1).Control(4)=   "txtDateScrew(0)"
         Tab(1).Control(5)=   "txtDateScrew(1)"
         Tab(1).Control(6)=   "Line1(7)"
         Tab(1).Control(7)=   "Line1(6)"
         Tab(1).Control(8)=   "Line1(5)"
         Tab(1).Control(9)=   "Line1(4)"
         Tab(1).Control(10)=   "Label1(7)"
         Tab(1).Control(11)=   "Label1(6)"
         Tab(1).Control(12)=   "Label1(5)"
         Tab(1).Control(13)=   "Label1(4)"
         Tab(1).ControlCount=   14
         TabCaption(2)   =   "Shiever Mesh 60"
         TabPicture(2)   =   "frmBlending.frx":688A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "GridShiever"
         Tab(2).Control(1)=   "txtShiever(4)"
         Tab(2).Control(2)=   "txtShiever(6)"
         Tab(2).Control(3)=   "txtShiever(5)"
         Tab(2).Control(4)=   "txtShiever(3)"
         Tab(2).Control(5)=   "txtShiever(2)"
         Tab(2).Control(6)=   "Frame3"
         Tab(2).Control(7)=   "txtShiever(1)"
         Tab(2).Control(8)=   "txtShiever(0)"
         Tab(2).Control(9)=   "txtDateShiever(0)"
         Tab(2).Control(10)=   "txtDateShiever(1)"
         Tab(2).Control(11)=   "Label3"
         Tab(2).Control(12)=   "Line1(15)"
         Tab(2).Control(13)=   "Line1(14)"
         Tab(2).Control(14)=   "Line1(13)"
         Tab(2).Control(15)=   "Line1(12)"
         Tab(2).Control(16)=   "Label1(18)"
         Tab(2).Control(17)=   "Label1(17)"
         Tab(2).Control(18)=   "Label1(16)"
         Tab(2).Control(19)=   "Label1(15)"
         Tab(2).Control(20)=   "Label2"
         Tab(2).Control(21)=   "Label1(14)"
         Tab(2).Control(22)=   "Label1(13)"
         Tab(2).Control(23)=   "Label1(12)"
         Tab(2).Control(24)=   "Line1(11)"
         Tab(2).Control(25)=   "Line1(10)"
         Tab(2).Control(26)=   "Line1(9)"
         Tab(2).Control(27)=   "Line1(8)"
         Tab(2).Control(28)=   "Label1(11)"
         Tab(2).Control(29)=   "Label1(10)"
         Tab(2).Control(30)=   "Label1(9)"
         Tab(2).Control(31)=   "Label1(8)"
         Tab(2).ControlCount=   32
         Begin VB.TextBox txtBlending 
            Appearance      =   0  'Flat
            DataField       =   "lotno"
            Height          =   330
            Index           =   2
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   6
            Tag             =   "Blend"
            Top             =   1320
            Width           =   1740
         End
         Begin VB.TextBox txtBlending 
            Appearance      =   0  'Flat
            DataField       =   "OrderID"
            DataSource      =   "DDE"
            Height          =   330
            Index           =   4
            Left            =   1710
            TabIndex        =   2
            Top             =   630
            Visible         =   0   'False
            Width           =   1740
         End
         Begin VB.TextBox txtBlending 
            Appearance      =   0  'Flat
            DataField       =   "grup"
            Height          =   330
            Index           =   1
            Left            =   1710
            TabIndex        =   8
            Tag             =   "Blend"
            Top             =   1665
            Width           =   1740
         End
         Begin VB.TextBox txtBlending 
            Appearance      =   0  'Flat
            DataField       =   "noItem"
            Height          =   330
            Index           =   0
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   4
            Tag             =   "Blend"
            Top             =   975
            Width           =   1740
         End
         Begin VB.CommandButton cmdLink 
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   3450
            Picture         =   "frmBlending.frx":68A6
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1335
            Width           =   405
         End
         Begin VB.CommandButton cmdLink 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   3435
            Picture         =   "frmBlending.frx":6C30
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   983
            Width           =   405
         End
         Begin VB.CommandButton cmdLink 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   3435
            Picture         =   "frmBlending.frx":6FBA
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   638
            Visible         =   0   'False
            Width           =   405
         End
         Begin MSComCtl2.DTPicker viewDate 
            Height          =   390
            Left            =   -65745
            TabIndex        =   20
            Top             =   795
            Visible         =   0   'False
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   688
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   16580610
            CurrentDate     =   39541
         End
         Begin MSDataGridLib.DataGrid gridScrew 
            Height          =   4095
            Left            =   -74835
            TabIndex        =   21
            Top             =   1440
            Width           =   11340
            _ExtentX        =   20003
            _ExtentY        =   7223
            _Version        =   393216
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
               DataField       =   "jam"
               Caption         =   "Jam"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "Suhu"
               Caption         =   "Suhu"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
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
         Begin MSDataGridLib.DataGrid GridShiever 
            Height          =   2775
            Left            =   -74865
            TabIndex        =   57
            Top             =   2475
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   4895
            _Version        =   393216
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
            ColumnCount     =   11
            BeginProperty Column00 
               DataField       =   "NoBag"
               Caption         =   "No Bag"
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
               DataField       =   "Mesh100_L80"
               Caption         =   ">80 (%) *"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "Mesh100_P80"
               Caption         =   "80 (%) *"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "Mesh100_P100"
               Caption         =   "100 (%) *"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "Mesh100_Total"
               Caption         =   "Total (%)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "Mesh150_P100"
               Caption         =   "100 (%) **"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "Mesh150_P150"
               Caption         =   "150 (%) **"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "Mesh150_Total"
               Caption         =   "total (%)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column08 
               DataField       =   "kondisi_powder"
               Caption         =   "Kondisi Powder ***"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column09 
               DataField       =   "kuantitas"
               Caption         =   "Kuantitas (Kg)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column10 
               DataField       =   "Operator"
               Caption         =   "Operator"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "0"
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
                  Alignment       =   1
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
               EndProperty
               BeginProperty Column07 
                  Alignment       =   1
               EndProperty
               BeginProperty Column08 
               EndProperty
               BeginProperty Column09 
                  Alignment       =   1
               EndProperty
               BeginProperty Column10 
               EndProperty
            EndProperty
         End
         Begin VB.TextBox txtShiever 
            Appearance      =   0  'Flat
            DataField       =   "lotno"
            Height          =   315
            Index           =   4
            Left            =   -71445
            Locked          =   -1  'True
            TabIndex        =   30
            Tag             =   "Blend"
            Top             =   1905
            Width           =   2310
         End
         Begin VB.TextBox txtShiever 
            Appearance      =   0  'Flat
            DataField       =   "losess"
            Height          =   315
            Index           =   6
            Left            =   -66360
            TabIndex        =   32
            Tag             =   "Blend"
            Top             =   1905
            Width           =   900
         End
         Begin VB.TextBox txtShiever 
            Appearance      =   0  'Flat
            DataField       =   "kuantitas_sisa"
            Height          =   315
            Index           =   5
            Left            =   -66360
            TabIndex        =   31
            Tag             =   "Blend"
            Top             =   1560
            Width           =   900
         End
         Begin VB.TextBox txtShiever 
            Appearance      =   0  'Flat
            DataField       =   "hasil_powder_bag"
            Height          =   315
            Index           =   3
            Left            =   -70035
            TabIndex        =   29
            Tag             =   "Blend"
            Top             =   1560
            Width           =   900
         End
         Begin VB.TextBox txtShiever 
            Appearance      =   0  'Flat
            DataField       =   "hasil_powder"
            Height          =   315
            Index           =   2
            Left            =   -71445
            TabIndex        =   28
            Tag             =   "Blend"
            Top             =   1560
            Width           =   900
         End
         Begin VB.Frame Frame3 
            Caption         =   "Kondisi Magnet"
            Height          =   885
            Left            =   -74775
            TabIndex        =   48
            Top             =   1470
            Width           =   1725
            Begin VB.OptionButton OpMagnet 
               Caption         =   "Kotor"
               Height          =   255
               Index           =   1
               Left            =   180
               TabIndex        =   27
               Tag             =   "Blend"
               Top             =   555
               Width           =   1170
            End
            Begin VB.OptionButton OpMagnet 
               Caption         =   "Bersih"
               Height          =   255
               Index           =   0
               Left            =   195
               TabIndex        =   26
               Tag             =   "Blend"
               Top             =   255
               Value           =   -1  'True
               Width           =   1170
            End
         End
         Begin VB.TextBox txtShiever 
            Appearance      =   0  'Flat
            DataField       =   "grup"
            Height          =   315
            Index           =   1
            Left            =   -74010
            Locked          =   -1  'True
            TabIndex        =   23
            Tag             =   "Blend"
            Top             =   885
            Width           =   1395
         End
         Begin VB.TextBox txtShiever 
            Appearance      =   0  'Flat
            DataField       =   "lotno"
            Height          =   315
            Index           =   0
            Left            =   -74010
            Locked          =   -1  'True
            TabIndex        =   22
            Tag             =   "Blend"
            Top             =   525
            Width           =   1395
         End
         Begin VB.TextBox txtScrew 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "grup"
            Height          =   330
            Index           =   1
            Left            =   -74010
            Locked          =   -1  'True
            TabIndex        =   17
            Tag             =   "Blend"
            Top             =   885
            Width           =   1395
         End
         Begin VB.TextBox txtScrew 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "lotno"
            Height          =   330
            Index           =   0
            Left            =   -74010
            Locked          =   -1  'True
            TabIndex        =   16
            Tag             =   "Blend"
            Top             =   525
            Width           =   1395
         End
         Begin MSDataGridLib.DataGrid gridBlending 
            Height          =   4200
            Left            =   4050
            TabIndex        =   15
            Top             =   1335
            Width           =   7515
            _ExtentX        =   13256
            _ExtentY        =   7408
            _Version        =   393216
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
               DataField       =   "prelot"
               Caption         =   "Pre Lot No"
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
               DataField       =   "qty"
               Caption         =   "QTY (Kg)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
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
         Begin VB.Frame Frame2 
            Caption         =   "Kondisi Blender"
            Height          =   1140
            Left            =   300
            TabIndex        =   39
            Top             =   3135
            Width           =   3165
            Begin VB.OptionButton OpBlender 
               Caption         =   "Kotor"
               Height          =   330
               Index           =   1
               Left            =   240
               TabIndex        =   12
               Tag             =   "Blend"
               Top             =   615
               Width           =   1530
            End
            Begin VB.OptionButton OpBlender 
               Caption         =   "Bersih"
               Height          =   330
               Index           =   0
               Left            =   255
               TabIndex        =   11
               Tag             =   "Blend"
               Top             =   270
               Value           =   -1  'True
               Width           =   1530
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Mesh"
            Height          =   1050
            Left            =   285
            TabIndex        =   38
            Top             =   2070
            Width           =   3180
            Begin VB.OptionButton OpMesh 
               Caption         =   "Mesh 150"
               Height          =   315
               Index           =   1
               Left            =   225
               TabIndex        =   10
               Tag             =   "Blend"
               Top             =   555
               Width           =   1350
            End
            Begin VB.OptionButton OpMesh 
               Caption         =   "Mesh 100"
               Height          =   315
               Index           =   0
               Left            =   225
               TabIndex        =   9
               Tag             =   "Blend"
               Top             =   210
               Value           =   -1  'True
               Width           =   1350
            End
         End
         Begin MSComCtl2.DTPicker txtDateBlending 
            DataField       =   "tanggal_mulai_blending"
            Height          =   315
            Index           =   0
            Left            =   9090
            TabIndex        =   13
            Tag             =   "Blend"
            Top             =   525
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   16580611
            CurrentDate     =   39536
         End
         Begin MSComCtl2.DTPicker txtDateBlending 
            DataField       =   "tanggal_selesai_blending"
            Height          =   315
            Index           =   1
            Left            =   9090
            TabIndex        =   14
            Tag             =   "Blend"
            Top             =   885
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   16580611
            CurrentDate     =   39536
         End
         Begin MSComCtl2.DTPicker txtDateScrew 
            DataField       =   "tanggal_mulai_screw"
            Height          =   315
            Index           =   0
            Left            =   -68490
            TabIndex        =   18
            Tag             =   "Blend"
            Top             =   525
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   16580611
            CurrentDate     =   39536
         End
         Begin MSComCtl2.DTPicker txtDateScrew 
            DataField       =   "tanggal_selesai_screw"
            Height          =   315
            Index           =   1
            Left            =   -68490
            TabIndex        =   19
            Tag             =   "Blend"
            Top             =   885
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   16580611
            CurrentDate     =   39536
         End
         Begin MSComCtl2.DTPicker txtDateShiever 
            DataField       =   "tanggal_mulai_shiever"
            Height          =   315
            Index           =   0
            Left            =   -68490
            TabIndex        =   24
            Tag             =   "Blend"
            Top             =   525
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy hh:mm"
            Format          =   16580611
            CurrentDate     =   39536
         End
         Begin MSComCtl2.DTPicker txtDateShiever 
            DataField       =   "tanggal_selesai_shiever"
            Height          =   315
            Index           =   1
            Left            =   -68490
            TabIndex        =   25
            Tag             =   "Blend"
            Top             =   885
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy hh:mm"
            Format          =   16580611
            CurrentDate     =   39536
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   1560
            TabIndex        =   61
            Top             =   5130
            Width           =   1905
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Approved By"
            Height          =   255
            Index           =   21
            Left            =   195
            TabIndex        =   62
            Top             =   5190
            Width           =   930
         End
         Begin VB.Line Line1 
            Index           =   18
            X1              =   195
            X2              =   1770
            Y1              =   5445
            Y2              =   5445
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Lot No"
            Height          =   255
            Index           =   19
            Left            =   270
            TabIndex        =   60
            Top             =   1365
            Width           =   930
         End
         Begin VB.Line Line1 
            Index           =   17
            X1              =   270
            X2              =   1710
            Y1              =   1635
            Y2              =   1635
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Manufature Order"
            Height          =   255
            Index           =   20
            Left            =   270
            TabIndex        =   59
            Top             =   660
            Visible         =   0   'False
            Width           =   2190
         End
         Begin VB.Line Line1 
            Index           =   16
            Visible         =   0   'False
            X1              =   2700
            X2              =   270
            Y1              =   945
            Y2              =   945
         End
         Begin VB.Label Label3 
            Caption         =   "* Mesh 100              ** Mesh 150                  *** Kondisi powder setelah melewati magnet"
            Height          =   255
            Left            =   -74760
            TabIndex        =   58
            Top             =   5325
            Width           =   9660
         End
         Begin VB.Line Line1 
            Index           =   15
            X1              =   -68130
            X2              =   -66270
            Y1              =   2205
            Y2              =   2205
         End
         Begin VB.Line Line1 
            Index           =   14
            X1              =   -68160
            X2              =   -66255
            Y1              =   1860
            Y2              =   1860
         End
         Begin VB.Line Line1 
            Index           =   13
            X1              =   -72840
            X2              =   -71220
            Y1              =   2205
            Y2              =   2205
         End
         Begin VB.Line Line1 
            Index           =   12
            X1              =   -72855
            X2              =   -71235
            Y1              =   1860
            Y2              =   1860
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Kg"
            Height          =   255
            Index           =   18
            Left            =   -65340
            TabIndex        =   56
            Top             =   1980
            Width           =   435
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Kg"
            Height          =   255
            Index           =   17
            Left            =   -65340
            TabIndex        =   55
            Top             =   1620
            Width           =   435
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Bag"
            Height          =   255
            Index           =   16
            Left            =   -69030
            TabIndex        =   54
            Top             =   1620
            Width           =   435
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Kg"
            Height          =   255
            Index           =   15
            Left            =   -70425
            TabIndex        =   53
            Top             =   1620
            Width           =   435
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Losess"
            Height          =   300
            Left            =   -68130
            TabIndex        =   52
            Top             =   1980
            Width           =   1080
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Kuantitas Sisa Powder"
            Height          =   255
            Index           =   14
            Left            =   -68145
            TabIndex        =   51
            Top             =   1620
            Width           =   1845
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Residual Lot No"
            Height          =   255
            Index           =   13
            Left            =   -72840
            TabIndex        =   50
            Top             =   1980
            Width           =   1500
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Hasil Powder"
            Height          =   255
            Index           =   12
            Left            =   -72840
            TabIndex        =   49
            Top             =   1620
            Width           =   930
         End
         Begin VB.Line Line1 
            Index           =   11
            X1              =   -71145
            X2              =   -67215
            Y1              =   1185
            Y2              =   1185
         End
         Begin VB.Line Line1 
            Index           =   10
            X1              =   -71145
            X2              =   -67215
            Y1              =   825
            Y2              =   825
         End
         Begin VB.Line Line1 
            Index           =   9
            X1              =   -74685
            X2              =   -73260
            Y1              =   1185
            Y2              =   1185
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   -74685
            X2              =   -73245
            Y1              =   825
            Y2              =   825
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal && Waktu Selesai Shiever"
            Height          =   255
            Index           =   11
            Left            =   -71145
            TabIndex        =   47
            Top             =   945
            Width           =   2625
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal && Waktu Mulai Shiever"
            Height          =   255
            Index           =   10
            Left            =   -71145
            TabIndex        =   46
            Top             =   570
            Width           =   2625
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Grup"
            Height          =   255
            Index           =   9
            Left            =   -74670
            TabIndex        =   45
            Top             =   945
            Width           =   930
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Lot No"
            Height          =   255
            Index           =   8
            Left            =   -74670
            TabIndex        =   44
            Top             =   570
            Width           =   930
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   -71145
            X2              =   -67215
            Y1              =   1185
            Y2              =   1185
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   -71160
            X2              =   -67230
            Y1              =   825
            Y2              =   825
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   -74655
            X2              =   -73230
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   -74670
            X2              =   -73230
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal && Waktu Selesai Screw"
            Height          =   255
            Index           =   7
            Left            =   -71145
            TabIndex        =   43
            Top             =   945
            Width           =   2625
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal && Waktu Mulai Screw"
            Height          =   255
            Index           =   6
            Left            =   -71145
            TabIndex        =   42
            Top             =   570
            Width           =   2625
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Grup"
            Height          =   255
            Index           =   5
            Left            =   -74670
            TabIndex        =   41
            Top             =   945
            Width           =   930
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Lot No"
            Height          =   255
            Index           =   4
            Left            =   -74670
            TabIndex        =   40
            Top             =   570
            Width           =   930
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   6450
            X2              =   10380
            Y1              =   1185
            Y2              =   1185
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   6450
            X2              =   10380
            Y1              =   825
            Y2              =   825
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   270
            X2              =   1845
            Y1              =   1980
            Y2              =   1980
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   270
            X2              =   1710
            Y1              =   1290
            Y2              =   1290
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal && Waktu Selesai Blending"
            Height          =   255
            Index           =   3
            Left            =   6435
            TabIndex        =   37
            Top             =   945
            Width           =   2625
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal && Waktu Mulai Blending"
            Height          =   255
            Index           =   2
            Left            =   6435
            TabIndex        =   36
            Top             =   570
            Width           =   2625
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Grup"
            Height          =   255
            Index           =   1
            Left            =   270
            TabIndex        =   35
            Top             =   1725
            Width           =   930
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Produk"
            Height          =   255
            Index           =   0
            Left            =   270
            TabIndex        =   34
            Top             =   1020
            Width           =   930
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MYDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5880
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1005
      BindFormTAG     =   "Blend"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmBlending"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RsDetail As New DBQuick
Private RsScrew As New DBQuick
Private RsShiever As New DBQuick
Private MEdit As Boolean
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RsLookup As New DBQuick
Private RsReseter As New DBQuick
Private curCol As Integer
Private isNewData As Boolean

Private Sub ResetGudang()
On Error GoTo 2
   With RsReseter.DBRecordset
   If .Recordcount > 0 Then
      .MoveFirst
      While Not .EOF
        '*** reset status barang digudang *** (iki benjut lek listrik mati)'
         SendDataToServer "update [Inventory Tabel] set StockTmp =" & .Fields("qty") & ",LockFIFO=0,Qty_Out=0 where sl_no='" & .Fields("Prelot") & "'"
         .MoveNext
      Wend
    End If
   End With
Exit Sub
2:
MessageBox Err.Description, "frmblending:resetgudang" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub SimpanDetail()
   On Error GoTo xErr
   ResetGudang
   '*** Simpan data detail Blending ***
   With RsDetail.DBRecordset
      If SendDataToServer("delete from blending_detail where lotNo ='" & MyDDE.GetFieldByName("LotNo") & "'") Then
         .MoveFirst
         While Not .EOF
            SendDataToServer "insert into blending_detail (lotNo,prelot,qty) values ('" & MyDDE.GetFieldByName("Lotno") & "','" & _
                             .Fields("prelot") & "'," & FQty(.Fields("qty")) & ")"
                             
            '*** Update Status Residual
            SendDataToServer "update residual_lot set status = 1 where sl_no='" & .Fields("prelot") & "'"
            
            '*** Update Gudang
            SendDataToServer "update [Inventory Tabel] set StockTmp = 0, LockFIFO=1, Qty_Out=" & .Fields("qty") & " where sl_no='" & .Fields("Prelot") & "'"
            .MoveNext
         Wend
      End If
   End With
   
   '*** Simpan data sCREW ***'
   With RsScrew.DBRecordset
      If SendDataToServer("delete from Screw where lotNo ='" & MyDDE.GetFieldByName("LotNo") & "'") Then
         .MoveFirst
         While Not .EOF
            SendDataToServer "insert into Screw (lotNo,jam,suhu) values ('" & _
                             MyDDE.GetFieldByName("Lotno") & "','" & _
                             Format(.Fields("jam"), "yyyy-MM-dd hh:mm:ss") & "'," & FQty(.Fields("suhu")) & ")"
            .MoveNext
         Wend
      End If
   End With
   
   '*** Simpan data Shiever ***'
   With RsShiever.DBRecordset
      If SendDataToServer("delete from ShieverMesh60 where lotNo ='" & MyDDE.GetFieldByName("LotNo") & "'") Then
         .MoveFirst
         While Not .EOF
            SendDataToServer "insert into ShieverMesh60 (LotNo,NoBag,Mesh100_L80,Mesh100_P80,Mesh100_P100," & _
                                     "Mesh100_Total,Mesh150_P100,Mesh150_P150,Mesh150_Total,Kondisi_powder," & _
                                     "kuantitas,operator) values ('" & MyDDE.GetFieldByName("Lotno") & "','" & _
                                .Fields("NoBag") & "'," & FQty(.Fields("Mesh100_L80")) & "," & _
                                FQty(.Fields("Mesh100_P80")) & "," & FQty(.Fields("Mesh100_P100")) & "," & _
                                FQty(.Fields("Mesh100_Total")) & "," & FQty(.Fields("Mesh150_P100")) & "," & _
                                FQty(.Fields("Mesh150_P150")) & "," & FQty(.Fields("Mesh150_Total")) & ",'" & _
                                .Fields("kondisi_powder") & "'," & FQty(.Fields("Kuantitas")) & ",'" & _
                                .Fields("Operator") & "')"
            .MoveNext
         Wend
      End If
   End With
Exit Sub
xErr:
   MessageBox Err.Description & "at Saving Data", "frmBlending : SimpanDetail", msgOkOnly, msgExclamation
   Err.Clear
End Sub

Private Sub cmdLink_Click(Index As Integer)
On Error GoTo 1
   Select Case Index
      Case 0: RsLookup.DBOpen "select OrderID,OrderName,Type,RequireDate from [MAnufacture Order] where status='RELEASED'", CNN
      Case 1: RsLookup.DBOpen "Select noItem as Kode, ItemName as [Nama Produk] from Inventory where Manufacture = 1", CNN
      Case 2: RsLookup.DBOpen "Select sl_no,creation_date from item_tracking_list where sl_code='lot' and blocked=0", CNN
   End Select
   If RsLookup.DBRecordset.Recordcount > 0 Then
      Set mCall.FormData = RsLookup.DBRecordset
      Select Case Index
         Case 0: mCall.FromTagActive = "Manufacture Order"
         Case 1: mCall.FromTagActive = "PRODUK"
         Case 2: mCall.FromTagActive = "Lot No"
      End Select
   Else
      MessageBox "Data Tidak Tersedia", "Peringatan", msgOkOnly, msgCrtical
   End If
Exit Sub
1:
MessageBox Err.Description, "frmblending:cmdlink_click" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Form_Load()
   On Error GoTo 1
   With MyDDE
      Set .ActiveConnection = CNN
      Set .BindForm = Me
      .BindFormTAG = "Blend"
      .PrepareQuery = "select blending_header.*,inventory.itemName as itemName from blending_header left outer join inventory on blending_header.noItem=inventory.noItem where blending_header.status=0"
   End With
   HiasFormManTell Picture1, Me
   gridScrew.RowHeight = 300
   Set mCall = New frmCaller
Exit Sub
1:
MessageBox Err.Description, "frmblending:form_load" & Err.Number, msgOkOnly, msgExclamation

End Sub



Private Sub gridScrew_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo Hell
   curCol = gridScrew.col
   If MyDDE.ChildRecordset.Recordcount > 0 Then
      viewDate.Visible = False
      
      
      viewDate.Move gridScrew.Left + gridScrew.Columns(gridScrew.col).Left, _
                    gridScrew.Top + gridScrew.RowTop(gridScrew.row), _
                    gridScrew.Columns(gridScrew.col).width, _
                    gridScrew.RowHeight
      
      
      Select Case gridScrew.col
         Case 0:
            viewDate.Value = IIf(IsNull(MyDDE.ChildRecordset.Fields("jam")), Now, MyDDE.ChildRecordset.Fields("jam"))
            viewDate.Visible = True
      End Select
   End If
Exit Sub
Hell:
      If Err.Number = 6148 Then
         Err.Clear
      Else
         MessageBox Err.Description, "Error", msgOkOnly, msgCrtical
         Err.Clear
      End If
End Sub

Private Sub GridShiever_AfterColEdit(ByVal ColIndex As Integer)
   Select Case ColIndex
      Case 1, 2, 3
         MyDDE.ChildRecordset.Fields("Mesh100_Total") = Val(MyDDE.ChildRecordset.Fields("Mesh100_L80")) + _
                                                        Val(MyDDE.ChildRecordset.Fields("Mesh100_P80")) + _
                                                        Val(MyDDE.ChildRecordset.Fields("Mesh100_P100"))
      Case 5, 6
         MyDDE.ChildRecordset.Fields("Mesh150_Total") = Val(MyDDE.ChildRecordset.Fields("Mesh150_P100")) + _
                                                        Val(MyDDE.ChildRecordset.Fields("Mesh150_P150"))
   End Select
End Sub


Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
On Error GoTo 1
   Select Case UCase(mCall.FromTagActive)
      Case "PRE LOT NUMBER":
         MyDDE.ChildRecordset.MoveLast
         MyDDE.ChildRecordset.Fields("Prelot") = mCall.GetFieldByName("sl_no")
         MyDDE.ChildRecordset.Fields("qty") = mCall.GetFieldByName("qty_in")
      Case "LOT NO":
         MyDDE.GetFieldByName("lotno") = mCall.GetFieldByName(0)
         txtScrew(0).Text = mCall.GetFieldByName(0)
         txtShiever(0).Text = mCall.GetFieldByName(0)
         txtShiever(4).Text = mCall.GetFieldByName(0)
      Case "MANUFACTURE ORDER":
         txtBlending(4).Text = mCall.GetFieldByName(0)
      Case "PRODUK":
         MyDDE.GetFieldByName("NoItem") = mCall.GetFieldByName(0)
   End Select
Exit Sub
1:
MessageBox Err.Description, "frmblending:mcall_rowcolchange" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
   Dim strSQL1 As String
   Dim strSQL2 As String
   Select Case AdReasonActiveDb
      Case tmbAddNew
         MEdit = True
         isNewData = True
         MyDDE.GetFieldByName("Tanggal_mulai_blending") = Now
         MyDDE.GetFieldByName("Tanggal_selesai_blending") = Now
         MyDDE.GetFieldByName("Tanggal_mulai_Screw") = Now
         MyDDE.GetFieldByName("Tanggal_selesai_Screw") = Now
         MyDDE.GetFieldByName("Tanggal_mulai_Shiever") = Now
         MyDDE.GetFieldByName("Tanggal_selesai_shiever") = Now
         MyDDE.GetFieldByName("hasil_powder") = 0
         MyDDE.GetFieldByName("hasil_powder_bag") = 0
         MyDDE.GetFieldByName("residual") = "-"
         MyDDE.GetFieldByName("kuantitas_sisa") = 0
         MyDDE.GetFieldByName("losess") = 0
         txtDateBlending(0).Value = Now
         txtDateBlending(1).Value = Now
         txtDateScrew(0).Value = Now
         txtDateScrew(1).Value = Now
         txtDateShiever(0).Value = Now
         txtDateShiever(1).Value = Now
      Case tmbDetail
         Select Case SSTab1.Tab
            Case 0:
               strSQL1 = "Select sl_no,qty_in from blending_prelot " & SQLLookupParameter(MyDDE.ChildRecordset, "sl_no", "prelot", "(Manufacture = 2) and LockFIFO=0 and itemName='" & txtBlending(0).Text & "'")
               strSQL2 = "select sl_no,qty as qty_in from residual_lot where status=0 and noItem ='" & txtBlending(0).Text & "'"
               RsLookup.DBOpen strSQL1 & " union " & strSQL2, CNN
               If RsLookup.Recordcount > 0 Then
                  Set mCall.FormData = RsLookup.DBRecordset
                  mCall.FromTagActive = "Pre Lot Number"
               Else
                  MessageBox "Data Tidak Tersedia", "Peringatan", msgOkOnly, msgCrtical
                  MyDDE.ChildRecordset.MoveLast
                  MyDDE.ChildRecordset.Delete
               End If
            Case 1:
               MyDDE.ChildRecordset.Fields("Jam") = Now
               MyDDE.ChildRecordset.Fields("Suhu") = 0
            Case 2:
               MyDDE.ChildRecordset.Fields("Operator") = MainMenu.StatusBar1.Panels(1).Text
               MyDDE.ChildRecordset.Fields("Mesh100_L80") = 0
               MyDDE.ChildRecordset.Fields("Mesh100_P80") = 0
               MyDDE.ChildRecordset.Fields("Mesh100_P100") = 0
               MyDDE.ChildRecordset.Fields("Mesh100_Total") = 0
               MyDDE.ChildRecordset.Fields("Mesh150_P100") = 0
               MyDDE.ChildRecordset.Fields("Mesh150_P150") = 0
               MyDDE.ChildRecordset.Fields("Mesh150_Total") = 0
               MyDDE.ChildRecordset.Fields("kuantitas") = 0
         End Select
      Case tmbEdit
         MEdit = True
         isNewData = False
      Case tmbSave
         On Error GoTo xErr
         If MyDDE.IsChildMemberReady = True Then
            '*** BLock item trackking list***'
            SendDataToServer "update item_tracking_list set Blocked=1 where sl_no ='" & txtBlending(2).Text & "'"
            
            '*** Saving Resudual lot No ***'
            If (txtShiever(5).Text <> "0") And (isNewData = True) Then
               SendDataToServer "Insert into residual_lot (sl_no,qty,noItem) values ('Residual" & txtShiever(4).Text & "'," & FQty(txtShiever(5).Text) & ",'" & txtBlending(0).Text & "')"
            End If
            
            '*** Save Detail ***'
            SimpanDetail
         End If
         MEdit = False
      Case tmbCancel
         MEdit = False
         
      'Case tmbGrantAccess:
         
   End Select
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear
End Sub

Private Sub PrepareSQL()
On Error GoTo xErr
With MyDDE
   Dim strSQL As String
   strSQL = "insert into blending_header (LotNo,grup,Tanggal_mulai_blending,Tanggal_selesai_blending, " & _
                        "mesh,kondisi_blender,Tanggal_mulai_screw,tanggal_selesai_screw,tanggal_mulai_shiever," & _
                        "tanggal_selesai_shiever,kondisi_magnet,hasil_powder,hasil_powder_bag,residual, " & _
                        "kuantitas_sisa,losess,noItem,issued_by) values ('" & .GetFieldByName("LotNo") & "','" & _
                            .GetFieldByName("grup") & "','" & _
                            Format(txtDateBlending(0).Value, "yyyy-MM-dd hh:mm:ss") & "','" & _
                            Format(txtDateBlending(1).Value, "yyyy-MM-dd hh:mm:ss") & "','" & _
                            IIf(OpMesh(0).Value = True, "100", "150") & "','" & _
                            IIf(OpBlender(0).Value = True, "Bersih", "Kotor") & "','" & _
                            Format(txtDateScrew(0).Value, "yyyy-MM-dd hh:mm:ss") & "','" & _
                            Format(txtDateScrew(1).Value, "yyyy-MM-dd hh:mm:ss") & "','" & _
                            Format(txtDateShiever(0).Value, "yyyy-MM-dd hh:mm:ss") & "','" & _
                            Format(txtDateShiever(1).Value, "yyyy-MM-dd hh:mm:ss") & "','" & _
                            IIf(OpMagnet(0).Value = True, "Bersih", "Kotor") & "'," & _
                            FQty(.GetFieldByName("Hasil_powder")) & "," & _
                            FQty(.GetFieldByName("hasil_powder_bag")) & ",'" & _
                            .GetFieldByName("residual") & "'," & _
                            FQty(.GetFieldByName("kuantitas_sisa")) & "," & _
                            FQty(.GetFieldByName("losess")) & ",'" & _
                            txtBlending(0).Text & ",'" & MainMenu.StatusBar1.Panels(1).Text & "')"
    Debug.Print strSQL
                            
   .PrepareAppend = strSQL
   .PrepareUpdate = "update blending_header set grup ='" & .GetFieldByName("grup") & "'," & _
                           "Tanggal_mulai_blending ='" & Format(txtDateBlending(0).Value, "yyyy-MM-dd hh:mm:ss") & "'," & _
                           "Tanggal_selesai_blending ='" & Format(txtDateBlending(1).Value, "yyyy-MM-dd hh:mm:ss") & "', " & _
                           "mesh ='" & IIf(OpMesh(0).Value = True, "100", "150") & "'," & _
                           "kondisi_blender ='" & IIf(OpBlender(0).Value = True, "Bersih", "Kotor") & "'," & _
                           "Tanggal_mulai_screw ='" & Format(txtDateScrew(0).Value, "yyyy-MM-dd hh:mm:ss") & "'," & _
                           "tanggal_selesai_screw ='" & Format(txtDateScrew(1).Value, "yyyy-MM-dd hh:mm:ss") & "'," & _
                           "tanggal_mulai_shiever ='" & Format(txtDateShiever(0).Value, "yyyy-MM-dd hh:mm:ss") & "'," & _
                           "tanggal_selesai_shiever ='" & Format(txtDateShiever(1).Value, "yyyy-MM-dd hh:mm:ss") & "'," & _
                           "kondisi_magnet ='" & IIf(OpMagnet(0).Value = True, "Bersih", "Kotor") & "'," & _
                           "hasil_powder =" & FQty(.GetFieldByName("Hasil_powder")) & "," & _
                           "hasil_powder_bag =" & FQty(.GetFieldByName("hasil_powder_bag")) & "," & _
                           "residual ='" & .GetFieldByName("residual") & "', " & _
                           "kuantitas_sisa =" & FQty(.GetFieldByName("kuantitas_sisa")) & "," & _
                           "losess =" & FQty(.GetFieldByName("losess")) & "," & _
                           "noItem='" & txtBlending(0).Text & "' where LotNo ='" & .GetFieldByName("LotNo") & "'"
                           
   .PrepareDelete = "delete from blending_header where LotNo ='" & .GetFieldByName("LotNo") & "'"

End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 1
   PrepareSQL
   Select Case AdReasonActiveDb
      Case tmbDelete:
         If RsDetail.DBRecordset.Recordcount = 0 Then
            ResetGudang
         End If
   End Select
Exit Sub
1:
MessageBox Err.Description, "frmblending:mydde_executeorder" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   OpenDetail SSTab1.Tab
   Label4.Caption = IIf(IsNull(MyDDE.GetFieldByName("Approved_by")), "", MyDDE.GetFieldByName("Approved_by"))
End Sub

Private Sub OpenDetail(nTab As Integer)
On Error GoTo 1
   RsDetail.DBOpen "Select Prelot,qty from blending_detail where LotNo ='" & MyDDE.GetFieldByName("LotNo") & "'", CNN
   RsReseter.DBOpen "Select Prelot,qty from blending_detail where LotNo ='" & MyDDE.GetFieldByName("LotNo") & "'", CNN
   RsScrew.DBOpen "Select jam,suhu from Screw where LotNo ='" & MyDDE.GetFieldByName("LotNo") & "'", CNN
   RsShiever.DBOpen "Select NoBag,Mesh100_L80,Mesh100_P80,Mesh100_P100,Mesh100_Total,Mesh150_P100,Mesh150_P150,Mesh150_Total,Kondisi_powder,kuantitas,operator from ShieverMesh60 where LotNo ='" & MyDDE.GetFieldByName("LotNo") & "'", CNN
   Select Case nTab
      Case 0:
         Set MyDDE.ChildRecordset = RsDetail.DBRecordset
         Set gridBlending.DataSource = MyDDE.ChildRecordset
      Case 1:
         Set MyDDE.ChildRecordset = RsScrew.DBRecordset
         Set gridBlending.DataSource = MyDDE.ChildRecordset
      Case 2:
         Set MyDDE.ChildRecordset = RsShiever.DBRecordset
         Set gridBlending.DataSource = MyDDE.ChildRecordset
   End Select
   If MyDDE.GetFieldByName("mesh") = "100" Then
      OpMesh(0).Value = True
      OpMesh(1).Value = False
   Else
      OpMesh(0).Value = False
      OpMesh(1).Value = True
   End If
   If MyDDE.GetFieldByName("kondisi_blender") = "Bersih" Then
      OpBlender(0).Value = True
      OpBlender(1).Value = False
   Else
      OpBlender(0).Value = False
      OpBlender(1).Value = True
   End If
   If MyDDE.GetFieldByName("kondisi_magnet") = "Bersih" Then
      OpMagnet(0).Value = True
      OpMagnet(1).Value = False
   Else
      OpMagnet(0).Value = False
      OpMagnet(1).Value = True
   End If
Exit Sub
1:
MessageBox Err.Description, "frmblending:opendetail" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 2
   PrepareSQL
   cmdLink(0).Enabled = False
   cmdLink(1).Enabled = False
   cmdLink(2).Enabled = False
   Select Case AdReasonActiveDb
      Case tmbAddNew, tmbEdit, tmbDetail:
         cmdLink(0).Enabled = True
         cmdLink(1).Enabled = True
         cmdLink(2).Enabled = True
      Case tmbSave:
         MyDDE.IsChildMemberReady = True
      Case tmbDelete
   End Select
Exit Sub
2:
MessageBox Err.Description, "frmblending:mydde_prepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    On Error GoTo 1
   Select Case SSTab1.Tab
      Case 0:
         Set MyDDE.ChildRecordset = RsDetail.DBRecordset
         Set gridBlending.DataSource = MyDDE.ChildRecordset
      Case 1:
         Set MyDDE.ChildRecordset = RsScrew.DBRecordset
         Set gridScrew.DataSource = MyDDE.ChildRecordset
      Case 2:
         Set MyDDE.ChildRecordset = RsShiever.DBRecordset
         Set GridShiever.DataSource = MyDDE.ChildRecordset
   End Select
   Exit Sub
1:
MessageBox Err.Description, "frmblending:sstab1_click" & Err.Number, msgOkOnly, msgExclamation
End Sub


Private Sub txtBlending_Change(Index As Integer)
   txtScrew(1).Text = txtBlending(1).Text
   txtShiever(1).Text = txtBlending(1).Text
End Sub

Private Sub viewDate_Change()
   Select Case curCol
      Case 0: MyDDE.ChildRecordset.Fields("Jam") = viewDate.Value
   End Select
End Sub
