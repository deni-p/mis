VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmPlanned 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Planned Order"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPlanned.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   Tag             =   "Planned Order"
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      Height          =   6480
      Left            =   0
      ScaleHeight     =   6450
      ScaleWidth      =   11265
      TabIndex        =   5
      Top             =   0
      Width           =   11295
      Begin TabDlg.SSTab SSTab1 
         Height          =   6015
         Left            =   150
         TabIndex        =   0
         Top             =   165
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   10610
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         BackColor       =   15380335
         TabCaption(0)   =   "Create PO"
         TabPicture(0)   =   "FrmPlanned.frx":6852
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MonthView1(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "DGDetail(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdOk(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Create MO"
         TabPicture(1)   =   "FrmPlanned.frx":686E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdOk(1)"
         Tab(1).Control(1)=   "MonthView1(1)"
         Tab(1).Control(2)=   "DGDetail(1)"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "RPB"
         TabPicture(2)   =   "FrmPlanned.frx":688A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdOk(3)"
         Tab(2).Control(1)=   "DGDetail(2)"
         Tab(2).ControlCount=   2
         Begin VB.CommandButton cmdOk 
            BackColor       =   &H00EAAF6F&
            Caption         =   "Create RPB"
            Height          =   420
            Index           =   3
            Left            =   -65640
            TabIndex        =   13
            Top             =   5520
            Width           =   1575
         End
         Begin VB.CommandButton cmdOk 
            BackColor       =   &H00EAAF6F&
            Caption         =   "Create PO"
            Height          =   420
            Index           =   0
            Left            =   9360
            TabIndex        =   3
            Top             =   5520
            Width           =   1575
         End
         Begin VB.CommandButton cmdOk 
            BackColor       =   &H00EAAF6F&
            Caption         =   "Create MO"
            Height          =   420
            Index           =   1
            Left            =   -65640
            TabIndex        =   6
            Top             =   5520
            Width           =   1575
         End
         Begin MSDataGridLib.DataGrid DGDetail 
            Bindings        =   "FrmPlanned.frx":68A6
            Height          =   5070
            Index           =   0
            Left            =   60
            TabIndex        =   2
            Tag             =   "Partner"
            Top             =   390
            Width           =   10890
            _ExtentX        =   19209
            _ExtentY        =   8943
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            BackColor       =   16777215
            BorderStyle     =   0
            HeadLines       =   1
            RowHeight       =   15
            TabAction       =   1
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
            ColumnCount     =   13
            BeginProperty Column00 
               DataField       =   "Item ID"
               Caption         =   "Item ID"
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
               DataField       =   "Description"
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
               DataField       =   "Convert"
               Caption         =   "Transfer"
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
            BeginProperty Column03 
               DataField       =   "M_OR_P"
               Caption         =   "Type Item"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "M"
                  FalseValue      =   "P"
                  NullValue       =   "P"
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   7
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "Partner ID"
               Caption         =   "Partner ID"
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
               DataField       =   "Company"
               Caption         =   "Company"
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
               DataField       =   "Suggest QTY"
               Caption         =   "Suggest QTY"
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
            BeginProperty Column07 
               DataField       =   "Order QTY"
               Caption         =   "Order QTY"
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
               DataField       =   "Required Date"
               Caption         =   "Required Date"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd MMM yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column09 
               DataField       =   "Suggest Order Date"
               Caption         =   "Suggest Date"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd MMM yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column10 
               DataField       =   "Order Date"
               Caption         =   "Order Date"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd MMM yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column11 
               DataField       =   "OrderID"
               Caption         =   "OrderID"
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
               DataField       =   "Note"
               Caption         =   "Note"
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
               BeginProperty Column04 
               EndProperty
               BeginProperty Column05 
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
               EndProperty
               BeginProperty Column10 
               EndProperty
               BeginProperty Column11 
               EndProperty
               BeginProperty Column12 
               EndProperty
            EndProperty
         End
         Begin MSComCtl2.MonthView MonthView1 
            Height          =   2460
            Index           =   0
            Left            =   4575
            TabIndex        =   7
            Top             =   2610
            Visible         =   0   'False
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   4339
            _Version        =   393216
            ForeColor       =   -2147483630
            BackColor       =   -2147483633
            Appearance      =   0
            Enabled         =   0   'False
            StartOfWeek     =   16449538
            CurrentDate     =   38537
         End
         Begin MSComCtl2.MonthView MonthView1 
            Height          =   2460
            Index           =   1
            Left            =   -71235
            TabIndex        =   8
            Top             =   3015
            Visible         =   0   'False
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   4339
            _Version        =   393216
            ForeColor       =   -2147483630
            BackColor       =   -2147483633
            Appearance      =   0
            Enabled         =   0   'False
            StartOfWeek     =   16449538
            CurrentDate     =   38537
         End
         Begin MSDataGridLib.DataGrid DGDetail 
            Bindings        =   "FrmPlanned.frx":68BB
            Height          =   5070
            Index           =   1
            Left            =   -74925
            TabIndex        =   9
            Tag             =   "Partner"
            Top             =   390
            Width           =   10890
            _ExtentX        =   19209
            _ExtentY        =   8943
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            BackColor       =   16777215
            BorderStyle     =   0
            HeadLines       =   1
            RowHeight       =   15
            TabAction       =   1
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
            ColumnCount     =   13
            BeginProperty Column00 
               DataField       =   "Item ID"
               Caption         =   "Item ID"
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
               DataField       =   "Description"
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
               DataField       =   "Convert"
               Caption         =   "Transfer"
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
            BeginProperty Column03 
               DataField       =   "M_OR_P"
               Caption         =   "Type Item"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "M"
                  FalseValue      =   "P"
                  NullValue       =   "P"
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   7
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "Partner ID"
               Caption         =   "Partner ID"
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
               DataField       =   "Company"
               Caption         =   "Company"
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
               DataField       =   "Suggest QTY"
               Caption         =   "Suggest QTY"
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
            BeginProperty Column07 
               DataField       =   "Order QTY"
               Caption         =   "Order QTY"
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
               DataField       =   "Required Date"
               Caption         =   "Required Date"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd MMM yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column09 
               DataField       =   "Suggest Order Date"
               Caption         =   "Suggest Date"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd MMM yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column10 
               DataField       =   "Order Date"
               Caption         =   "Order Date"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd MMM yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column11 
               DataField       =   "OrderID"
               Caption         =   "OrderID"
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
               DataField       =   "Note"
               Caption         =   "Note"
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
               BeginProperty Column04 
               EndProperty
               BeginProperty Column05 
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
               EndProperty
               BeginProperty Column10 
               EndProperty
               BeginProperty Column11 
               EndProperty
               BeginProperty Column12 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid DGDetail 
            Bindings        =   "FrmPlanned.frx":68D0
            Height          =   5070
            Index           =   2
            Left            =   -74925
            TabIndex        =   12
            Tag             =   "Partner"
            Top             =   390
            Width           =   10890
            _ExtentX        =   19209
            _ExtentY        =   8943
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            BackColor       =   16777215
            BorderStyle     =   0
            HeadLines       =   1
            RowHeight       =   15
            TabAction       =   1
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
            ColumnCount     =   13
            BeginProperty Column00 
               DataField       =   "Item ID"
               Caption         =   "Item ID"
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
               DataField       =   "Description"
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
               DataField       =   "Convert"
               Caption         =   "Transfer"
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
            BeginProperty Column03 
               DataField       =   "M_OR_P"
               Caption         =   "Type Item"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "M"
                  FalseValue      =   "P"
                  NullValue       =   "P"
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   7
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "Partner ID"
               Caption         =   "Partner ID"
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
               DataField       =   "Company"
               Caption         =   "Company"
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
               DataField       =   "Suggest QTY"
               Caption         =   "Suggest QTY"
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
            BeginProperty Column07 
               DataField       =   "Order QTY"
               Caption         =   "Order QTY"
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
               DataField       =   "Required Date"
               Caption         =   "Required Date"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd MMM yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column09 
               DataField       =   "Suggest Order Date"
               Caption         =   "Suggest Date"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd MMM yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column10 
               DataField       =   "Order Date"
               Caption         =   "Order Date"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd MMM yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column11 
               DataField       =   "OrderID"
               Caption         =   "OrderID"
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
               DataField       =   "Note"
               Caption         =   "Note"
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
               BeginProperty Column04 
               EndProperty
               BeginProperty Column05 
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
               EndProperty
               BeginProperty Column10 
               EndProperty
               BeginProperty Column11 
               EndProperty
               BeginProperty Column12 
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   11265
      TabIndex        =   10
      Top             =   6465
      Width           =   11295
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Close"
         Height          =   420
         Index           =   2
         Left            =   9555
         TabIndex        =   4
         Top             =   135
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   330
         Left            =   555
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   165
         Width           =   4185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter"
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   210
         Width           =   390
      End
   End
End
Attribute VB_Name = "FrmPlanned"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RcPlan As New DBQuick


Private Sub cmdOk_Click(Index As Integer)
Dim I As Integer
Dim IDGen As New IDGenerator
Dim newPO As String

Select Case Index
       Case 0, 1:
           If RcPlan.DBRecordset.Recordcount <> 0 Then
                I = MessageBox("Anda yakin untuk melakukan transaksi?", "Konfirmasi", msgYesNo, msgQuestion)
                Screen.MousePointer = vbHourglass
                If I = 1 Then
                   If Index = 0 Then
                      TranferPO
                   Else
                      TranferJob
                   End If
                   SendDataToServer ("DELETE FROM [Planned Order] WHERE     ([Convert] = 1)")
                   Screen.MousePointer = vbDefault
                   MessageBox "Proses selesai.", "Informasi", msgOkOnly, msgInfo
                End If
                Call SSTab1_Click(SSTab1.Tab)
           Else
               MessageBox "No Suggest Order to create Manufacturing Order" & Chr(13) & "Choose option on transfer column", "Warning", msgOkOnly, msgCrtical
           End If
       Case 2: Unload Me
       Case 3:
         If RcPlan.DBRecordset.Recordcount > 0 Then
           newPO = IDGen.GetID("PO")
           
           '*** Saving HEader
           SendDataToServer " INSERT INTO  [PO Order] ( [REquire Date]                        ,PurchaseID               ,EmpID                                        , PartnerID                                   ,DatePurchase                                   , TermPayment      ,  Periode           , TypeTrans    , blanked_date ) " & _
                        " VALUES ('" & Format(RcPlan.DBRecordset.Fields("Required date"), "yyyy-MM-dd") & "',N'" & newPO & "',N'" & MainMenu.StatusBar1.Panels(1).Text & "', N'" & RcPlan.DBRecordset.Fields("PartnerID") & "','" & Format(Now, "yyyy-MM-dd") & "', " & 30 & ", " & Val(Month(Now)) & ", N'BLANKED','" & Format(Now, "yyyy-MM-dd") & "')"
           
           '*** Saving Detail
           SendDataToServer " INSERT INTO [Detail PO] ( PurchaseID, NoItem, QTYPO, ItemSupplierID, POPrice, ScheduleDate, VAT,QtyTemp,Hpp,sppid,tipe_item,curID,rate)" & _
                        " VALUES (N'" & newPO & "', N'" & RcPlan.DBRecordset.Fields("NoItem") & "', " & FQty(RcPlan.DBRecordset.Fields("Suggest QTY")) & ", N'" & RcPlan.DBRecordset.Fields("uom") & "', " & 0 & ", convert(Datetime,'" & Format(RcPlan.DBRecordset.Fields("Suggest Order Date"), "dd/mm/yy") & "',3), " & CDbl(RcPlan.DBRecordset.Fields("VAT")) & ", " & FQty(RcPlan.DBRecordset.Fields("Suggest QTYP")) & "," & "0" & ",'','" & "I" & "','" & "IDR" & "'," & "1" & ")"
         Else
            MessageBox "Data Kosong", "Peringatan", msgOkOnly, msgCrtical
         End If


End Select
End Sub

Private Sub Combo1_Click()
If Combo1.Text = "All" Then
'   RcPlan.DBOpen "SELECT  [Planned Order].NoItem AS [Item ID], [Planned Order].[DESC] AS Description, [Planned Order].[Convert], [Planned Order].M_OR_P,                       [Planned Order].PartnerID AS [Partner ID], PartnerDB.CompanyName AS Company, [Planned Order].[Suggest QTY] AS [Suggest QTY],                       [Planned Order].[Order QTY] AS [Order QTY], [Planned Order].[Required Date], [Planned Order].[Suggest Order Date], [Planned Order].[Order Date] FROM         [Planned Order] INNER JOIN PartnerDB ON [Planned Order].PartnerID = PartnerDB.PartnerID ORDER BY [Planned Order].PartnerID, [Planned Order].NoItem", Cnn, lckLockBatch
'   Set DGDetail(SSTab1.Tab).DataSource = RcPlan.DBRecordset
'   DGDetail(SSTab1.Tab).Columns(2).Button = True
'   DGDetail(SSTab1.Tab).Columns(3).Button = True
'Check1_Click
Call SSTab1_Click(SSTab1.Tab)
Else
   RcPlan.MakeFilter "[" & RcPlan.DBRecordset.Fields(Dgdetail(SSTab1.Tab).col).Name & "]  = '" & Combo1.Text & "'"
End If
End Sub

Private Sub Combo1_DropDown()
CreateFilter
End Sub

Private Sub DGDETAIL_ButtonClick(Index As Integer, ByVal ColIndex As Integer)
Select Case ColIndex
       Case 2:
            If RcPlan.DBRecordset.Recordcount <> 0 Then
               If RcPlan.DBRecordset.Fields("Convert") = True Then
                  RcPlan.DBRecordset.Fields("Convert") = False
               Else
                  RcPlan.DBRecordset.Fields("Convert") = True
               End If
               Dgdetail(SSTab1.Tab).Columns(2).Value = RcPlan.DBRecordset.Fields("Convert")
               SendDataToServer ("UPDATE    [Planned Order]" & _
                                 " Set [Convert] = " & BoolToInt(RcPlan.DBRecordset.Fields("Convert")) & _
                                 " WHERE     (ID = '" & RcPlan.DBRecordset.Fields("ID") & "')")
            End If
       Case 10: MoveControl
       
End Select
End Sub

Private Sub DGDETAIL_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
If Dgdetail(SSTab1.Tab).col = 2 Or Dgdetail(SSTab1.Tab).col = 3 Or Dgdetail(SSTab1.Tab).col = 10 Then
   Dgdetail(SSTab1.Tab).MarqueeStyle = dbgFloatingEditor
Else
   Dgdetail(SSTab1.Tab).MarqueeStyle = dbgHighlightRow
End If
Select Case Dgdetail(SSTab1.Tab).col
       Case 2:
       Case 10:
End Select
If Dgdetail(SSTab1.Tab).col <= 0 Then Dgdetail(SSTab1.Tab).col = 0
Label1 = "Filter By " & Dgdetail(SSTab1.Tab).Columns(Dgdetail(SSTab1.Tab).col).Caption
End Sub

Private Sub Form_Load()
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
GridLayout
Call SSTab1_Click(0)
SSTab1.Tab = 0
'RcPlan.DBOpen "SELECT     [Planned Order].NoItem AS [Item ID], [Planned Order].[DESC] AS Description, [Planned Order].[Convert], [Planned Order].M_OR_P,                       [Planned Order].PartnerID AS [Partner ID], PartnerDB.CompanyName AS Company, [Planned Order].[Suggest QTY] AS [Suggest QTY],                       [Planned Order].[Order QTY] AS [Order QTY], [Planned Order].[Required Date], [Planned Order].[Suggest Order Date], [Planned Order].[Order Date],[Planned Order].[ID] FROM         [Planned Order] INNER JOIN PartnerDB ON [Planned Order].PartnerID = PartnerDB.PartnerID ORDER BY [Planned Order].PartnerID, [Planned Order].NoItem", Cnn, lckLockBatch
'Set DGDetail.DataSource = RcPlan.DBRecordset
'DGDetail.Columns(2).Button = True
'DGDetail.Columns(3).Button = True
'DGDetail.Columns(10).Button = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmPlanned = Nothing
End Sub

Private Sub MonthView1_DateDblClick(Index As Integer, ByVal DateDblClicked As Date)
On Error Resume Next
If RcPlan.DBRecordset.Recordcount <> 0 Then
    If MonthView1(SSTab1.Tab).Value < RcPlan.DBRecordset.Fields("Suggest Order Date") Then
       MonthView1(SSTab1.Tab).Enabled = False
       MonthView1(SSTab1.Tab).Visible = False
       RcPlan.DBRecordset.Fields("Order Date") = MonthView1(SSTab1.Tab).Value
    Else
       MonthView1(SSTab1.Tab).Value = RcPlan.DBRecordset.Fields("Order Date")
    End If
End If
Dgdetail(SSTab1.Tab).SetFocus
Err.Clear
End Sub

Private Sub MonthView1_LostFocus(Index As Integer)
MonthView1(0).Enabled = False
MonthView1(0).Visible = False
MonthView1(1).Enabled = False
MonthView1(1).Visible = False
Dgdetail(SSTab1.Tab).SetFocus
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Picture1.Parent.hwnd
End Sub

Private Sub CreateFilter()
Dim Rc As New DBQuick
Dim Avdata As Variant
Dim I As Integer
Dim StrQuery As String
Dim StrField As String
If Dgdetail(SSTab1.Tab).col <= 0 Then Dgdetail(SSTab1.Tab).col = 0
Select Case UCase(RcPlan.DBRecordset.Fields(Dgdetail(SSTab1.Tab).col).Name)
       Case "ITEM ID":
            StrField = "NoItem"
            StrQuery = " SELECT [Planned Order].[" & StrField & "] FROM [Planned Order] INNER JOIN" & _
                       " PartnerDB ON [Planned Order].PartnerID = PartnerDB.PartnerID GROUP BY [Planned Order].[" & StrField & "] ORDER BY [Planned Order].[" & StrField & "]"
       
       Case "DESCRIPTION":
            StrField = "DESC"
            StrQuery = " SELECT [Planned Order].[" & StrField & "] FROM [Planned Order] INNER JOIN" & _
                       " PartnerDB ON [Planned Order].PartnerID = PartnerDB.PartnerID GROUP BY [Planned Order].[" & StrField & "] ORDER BY [Planned Order].[" & StrField & "]"
       
       Case "PARTNER ID":
            StrField = "PartnerID"
            StrQuery = " SELECT [PartnerDB].[" & StrField & "] FROM [Planned Order] INNER JOIN" & _
                       " PartnerDB ON [Planned Order].PartnerID = PartnerDB.PartnerID GROUP BY [PartnerDB].[" & StrField & "] ORDER BY [PartnerDB].[" & StrField & "]"
            
       Case "COMPANY":
            StrField = "CompanyName"
            StrQuery = " SELECT [PartnerDB].[" & StrField & "] FROM [Planned Order] INNER JOIN" & _
                       " PartnerDB ON [Planned Order].PartnerID = PartnerDB.PartnerID GROUP BY [PartnerDB].[" & StrField & "] ORDER BY [PartnerDB].[" & StrField & "]"
       
       Case Else:
            StrField = RcPlan.DBRecordset.Fields(Dgdetail(SSTab1.Tab).col).Name
            StrQuery = " SELECT [Planned Order].[" & StrField & "] FROM [Planned Order] INNER JOIN" & _
                       " PartnerDB ON [Planned Order].PartnerID = PartnerDB.PartnerID GROUP BY [Planned Order].[" & StrField & "] ORDER BY [Planned Order].[" & StrField & "]"

End Select
'StrQuery = " SELECT [Planned Order].[" & StrField & "] FROM [Planned Order] INNER JOIN" & _
          " PartnerDB ON [Planned Order].PartnerID = PartnerDB.PartnerID GROUP BY [Planned Order].[" & StrField & "] ORDER BY [Planned Order].[" & StrField & "]"
         ' MsgBox StrQuery
'Debug.Print StrQuery
Rc.DBOpen StrQuery, CNN, lckLockReadOnly
Combo1.Clear
Combo1.AddItem "All"
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        If .Recordcount <> 0 Then
            Avdata = .Getrows(.Recordcount, adBookmarkFirst)
            For I = 0 To UBound(Avdata, 2)
                If Not IsNull(Avdata(0, I)) Then Combo1.AddItem Avdata(0, I)
            Next
        End If
     End If
     If Combo1.ListCount <> 0 Then Combo1.ListIndex = 0
     Rc.CloseDB
End With
Set Rc = Nothing
End Sub

Private Sub MoveControl()
MonthView1(SSTab1.Tab).Enabled = True
MonthView1(SSTab1.Tab).Visible = True
With Dgdetail(SSTab1.Tab)
     If .row >= 10 Then
        MonthView1(SSTab1.Tab).Move (.Columns(10).Left + .Columns(10).width) - (MonthView1(SSTab1.Tab).width - 60), (.RowTop(.row) + (.RowHeight * 2) - 400) - MonthView1(SSTab1.Tab).Height
     Else
        MonthView1(SSTab1.Tab).Move (.Columns(10).Left + .Columns(10).width) - (MonthView1(SSTab1.Tab).width - 60), (.RowTop(.row) + .RowHeight + 400)
     End If
     MonthView1(SSTab1.Tab).ZOrder (0)
     MonthView1(SSTab1.Tab).SetFocus
End With
End Sub

Private Sub TranferPO()
Dim MyData As New clsTransaksi
Dim RcHeader As New DBQuick
Dim RcDetail As New DBQuick
Dim NoPo As String
RcHeader.DBOpen "SELECT PartnerID, [Order Date] FROM [Planned Order] WHERE ([Convert] = 1) " & _
" GROUP BY PartnerID, [Order Date] ORDER BY [Order Date], PartnerID", CNN, lckLockReadOnly
With RcHeader.DBRecordset
     If .Recordcount <> 0 Then
        .MoveFirst
        Do
          If .EOF Then Exit Do
          RcDetail.DBOpen " SELECT [Planned Order].PartnerID, [Planned Order].[Order Date], [Planned Order].NoItem, " & _
          " [Planned Order].[Suggest QTY], MIN(ISNULL([Inventory Tabel].HPP, 0)) AS HPP FROM  " & _
          " [Planned Order] LEFT OUTER JOIN [Inventory Tabel] ON [Planned Order].NoItem = [Inventory Tabel].NoItem " & _
          "  WHERE ([Planned Order].[Convert] = 1) GROUP BY [Planned Order].PartnerID, " & _
          " [Planned Order].[Order Date], [Planned Order].NoItem, [Planned Order].[Suggest QTY] " & _
          " HAVING ([Planned Order].PartnerID = N'" & .Fields("PartnerID") & "') AND " & _
          " ([Planned Order].[Order Date] = CONVERT(DATETIME, '" & Format(.Fields("Order Date"), "dd/mm/yy") & "', 3)) " & _
          " ORDER BY [Planned Order].[Order Date], [Planned Order].PartnerID", CNN, lckLockReadOnly
          
          Dim IDGen As New IDGenerator
          NoPo = IDGen.GetID("PO")     'MyData.PrepareIndex(tmbTransaksiPO, 5, "1", TglIndex)
          Set IDGen = Nothing
          
          SendDataToServer " INSERT INTO  [PO Order] ( PurchaseID,EmpID, PartnerID,  " & _
          " DatePurchase,   Periode, TypeTrans) VALUES (N'" & NoPo & "',N'" & MainMenu.StatusBar1.Panels(1).Text & _
          "', N'" & .Fields("PartnerID") & "',convert(Datetime, '" & Format(.Fields("Order date"), "dd/mm/yy") & "',3) , " & mVarPeriode & ", N'ORDER' )"
          
          RcDetail.DBRecordset.MoveFirst
          Do
             If RcDetail.DBRecordset.EOF Then Exit Do
                SendDataToServer " INSERT INTO [Detail PO] ( PurchaseID, NoItem, QTYPO,  POPrice, ScheduleDate, VAT,QtyTemp,Hpp)" & _
                                 " VALUES (N'" & NoPo & "', N'" & RcDetail.DBRecordset.Fields("NoItem") & "', " & RcDetail.DBRecordset.Fields("Suggest QTY") & ", " & CDbl(RcDetail.DBRecordset.Fields("HPP")) & ", convert(Datetime,'" & Format(dDateBegin, "dd/mm/yy") & "',3), 0, " & RcDetail.DBRecordset.Fields("Suggest QTY") & "," & CDbl(RcDetail.DBRecordset.Fields("HPP")) & "  )"
             RcDetail.DBRecordset.MoveNext
          Loop
          .MoveNext
        Loop
        .MoveLast
        
     End If
End With
Set MyData = Nothing
End Sub

Private Sub TranferJob()
Dim MyData As New clsTransaksi
Dim RcHeader As New DBQuick
Dim RcDetail As New DBQuick
Dim NoPo As String
Dim I As Integer
RcHeader.DBOpen "SELECT  NoItem, [DESC], [Order Date], [Required Date], [Order QTY],PARTNERID FROM         [Planned Order] WHERE     ([Convert] = 1) ORDER BY NoItem, [Order Date]", CNN, lckLockReadOnly

With RcHeader.DBRecordset
     If .Recordcount <> 0 Then
        .MoveFirst
        Do
          If .EOF Then Exit Do
          NoPo = IndexAuto
          'Copy Stage
          SendDataToServer " INSERT INTO [Manufacture Order]" & _
                           " (EmpID,Noitem,DateIssued,OrderID, [QTY Order],PartnerID, OrderName, Type, Status, Note, [CreateDate], [Priority], [RequireDate], [EarliesDate], [StartDate], [FinishedDate])" & _
                           " VALUES (N'" & mVarUserID & "',N'" & .Fields("NoItem") & "',convert(Datetime,'" & Format(Date, "dd/mm/yy") & "',3), N'" & NoPo & "'," & CDbl(.Fields("Order QTY")) & ", N'" & .Fields("PARTNERID") & "', N'" & .Fields("DESC") & "', N'ASSEMBLE ORDER', N'QUOTED', N'Transfer from Planned order',convert(datetime,'" & Format(.Fields("Required Date"), "dd/mm/yy") & "',3),N'NORMAL',convert(datetime,'" & Format(.Fields("Order Date"), "dd/mm/yy") & "',3),convert(datetime,'" & Format(.Fields("Order Date"), "dd/mm/yy") & "',3),convert(datetime,'" & Format(.Fields("Order Date"), "dd/mm/yy") & "',3),convert(datetime,'" & Format(.Fields("Order Date"), "dd/mm/yy") & "',3))"
          
          RcDetail.DBOpen " SELECT  [BOM Stage Detail].NoLine, [BOM Stage Detail].SeqStageID, [BOM Stage Detail].Description AS Keterangan, [BOM Stage Detail].ResourcesID, [Resources Table].Description AS Resources, [BOM Stage Detail].StageNote AS Catatan FROM         [BOM Stage Detail] LEFT OUTER JOIN                       [Resources Table] ON [BOM Stage Detail].ResourcesID = [Resources Table].ResourcesID WHERE     ([BOM Stage Detail].NoItem = N'" & .Fields("NoItem") & "') ORDER BY [BOM Stage Detail].NoLine, [BOM Stage Detail].SeqStageID", CNN, lckLockReadOnly
          If RcDetail.DBRecordset.Recordcount <> 0 Then
                RcDetail.DBRecordset.MoveFirst
                Do
                   If RcDetail.DBRecordset.EOF Then Exit Do
                      SendDataToServer " INSERT INTO [Order Output Detail]" & _
                                       " (EndDate,StartDate,OrderID, SeqNo, StageID,  ResourcesID,Status)" & _
                                       " VALUES  (Convert(Datetime,'" & Format(.Fields("Order Date"), "dd/mm/yy HH:MM:SS") & "',3),Convert(Datetime,'" & Format(.Fields("Order Date"), "dd/mm/yy HH:MM:SS") & "',3),N'" & NoPo & "', " & RcDetail.DBRecordset.Fields("NoLine") & ", N'" & RcDetail.DBRecordset.Fields("SeqStageID") & "', N'" & RcDetail.DBRecordset.Fields("ResourcesID") & "',0)"
                   RcDetail.DBRecordset.MoveNext
                Loop
          End If
          'Stage Selesai
          'Copy Detail
          RcDetail.DBOpen " SELECT [BOM Stage Detail].NoLine,[BOM Component Detail].SeqStageID, [BOM Component Detail].Component AS [Kode Barang], Inventory.ItemName AS [Nama Komponen],[BOM Component Detail].UOM, Inventory.PartnerID , PartnerDB.CompanyName AS [Nama Perusahaan],[BOM Component Detail].QTYUsage FROM [BOM Component Detail] INNER JOIN" & _
                          " [BOM Stage Detail] ON [BOM Component Detail].SeqStageID = [BOM Stage Detail].SeqStageID AND  [BOM Component Detail].NoItem = [BOM Stage Detail].NoItem INNER JOIN Inventory INNER JOIN  PartnerDB ON Inventory.PartnerID = PartnerDB.PartnerID ON [BOM Component Detail].Component = Inventory.NoItem WHERE     ([BOM Component Detail].NoItem = N'" & .Fields("NoItem") & "') GROUP BY [BOM Component Detail].SeqStageID, [BOM Component Detail].Component, Inventory.ItemName, [BOM Component Detail].UOM, Inventory.PartnerID,                        PartnerDB.CompanyName, [BOM Component Detail].QTYUsage, [BOM Stage Detail].NoLine ORDER BY [BOM Stage Detail].NoLine, [BOM Component Detail].SeqStageID", CNN, lckLockReadOnly
          If RcDetail.DBRecordset.Recordcount <> 0 Then
                RcDetail.DBRecordset.MoveFirst
                I = 0
                Do
                   I = I + 1
                   If RcDetail.DBRecordset.EOF Then Exit Do
                      SendDataToServer (" INSERT INTO [Ord Comp Detail]" & _
                                        " (SeqNo, StageID, OrderID, NoItem, [DESC], UOM, [Quote Qty], [Actual Qty], Phantom, Complete, PartnerID)" & _
                                        " VALUES    (" & I & ", N'" & RcDetail.DBRecordset.Fields("SeqStageID") & "', N'" & NoPo & "', N'" & RcDetail.DBRecordset.Fields("Kode Barang") & "', N'" & RcDetail.DBRecordset.Fields("Nama Komponen") & "', N'" & RcDetail.DBRecordset.Fields("UOM") & "', " & CDbl(RcDetail.DBRecordset.Fields("qtyusage")) & ", 0, 0, 0, N'" & RcDetail.DBRecordset.Fields("PartnerID") & "')")
                   RcDetail.DBRecordset.MoveNext
                Loop
          End If
          'Detail Selesai
          .MoveNext
        Loop
        .MoveLast
        
     End If
End With
Set MyData = Nothing
End Sub


Private Function TglIndex() As String
Dim TglHari As String
Dim TglBulan As String
Dim TglTahun As String
TglIndex = "PO-" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End Function

Private Function IndexAuto() As String
Dim Rc As New DBQuick
Dim TglSaiki As String
Dim Inom As Long
TglSaiki = Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2)
Rc.DBOpen "SELECT MAX(RIGHT(OrderID, 5)) AS MaxNom FROM [Manufacture Order]         [Manufacture Order] WHERE     (GETDATE() = { fn NOW() })", CNN, lckLockReadOnly
With Rc
     If .DBRecordset.Recordcount <> 0 Then
        Inom = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + 1
     Else
        Inom = 1
     End If
     Select Case Len(Trim(Str(Inom)))
            Case 0: IndexAuto = "OD/" & TglSaiki & "-" & Trim(Str(Inom))
            Case 1: IndexAuto = "OD/" & TglSaiki & "-" & "0000" & Trim(Str(Inom))
            Case 2: IndexAuto = "OD/" & TglSaiki & "-" & "000" & Trim(Str(Inom))
            Case 3: IndexAuto = "OD/" & TglSaiki & "-" & "00" & Trim(Str(Inom))
            Case 4: IndexAuto = "OD/" & TglSaiki & "-" & "0" & Trim(Str(Inom))
     End Select
End With
End Function

Private Sub GridLayout()
Dgdetail(0).Columns(0).width = 1800
Dgdetail(0).Columns(1).width = 1950.236
Dgdetail(0).Columns(2).width = 1019.906
Dgdetail(0).Columns(3).width = 1019.906
Dgdetail(0).Columns(4).width = 1514.835
Dgdetail(0).Columns(5).width = 1514.835
Dgdetail(0).Columns(6).width = 1514.835
Dgdetail(0).Columns(7).width = 1514.835
Dgdetail(0).Columns(8).width = 1514.835
Dgdetail(0).Columns(9).width = 1514.835
Dgdetail(0).Columns(10).width = 1514.835
Dgdetail(0).Columns(11).width = 1785.26
Dgdetail(0).Columns(12).width = 4020.095

Dgdetail(1).Columns(0).width = 1800
Dgdetail(1).Columns(1).width = 1950.236
Dgdetail(1).Columns(2).width = 1019.906
Dgdetail(1).Columns(3).width = 1019.906
Dgdetail(1).Columns(4).width = 1514.835
Dgdetail(1).Columns(5).width = 1514.835
Dgdetail(1).Columns(6).width = 1514.835
Dgdetail(1).Columns(7).width = 1514.835
Dgdetail(1).Columns(8).width = 1514.835
Dgdetail(1).Columns(9).width = 1514.835
Dgdetail(1).Columns(10).width = 1514.835
Dgdetail(1).Columns(11).width = 1785.26
Dgdetail(1).Columns(12).width = 4020.095
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 1 Then
   SSTab1.Caption = "Create MO"
   RcPlan.DBOpen "SELECT     [Planned Order].NoItem AS [Item ID], [Planned Order].[DESC] AS Description, [Planned Order].[Convert], [Planned Order].M_OR_P,                       [Planned Order].PartnerID AS [Partner ID], PartnerDB.CompanyName AS Company, [Planned Order].[Suggest QTY] AS [Suggest QTY],                       [Planned Order].[Order QTY] AS [Order QTY], [Planned Order].[Required Date], [Planned Order].[Suggest Order Date], [Planned Order].[Order Date],[Planned Order].[ID],[Planned Order].OrderID, [Planned Order].Note FROM         [Planned Order] INNER JOIN PartnerDB ON [Planned Order].PartnerID = PartnerDB.PartnerID where [Planned Order].M_OR_P = 1 ORDER BY [Planned Order].PartnerID, [Planned Order].NoItem", CNN, lckLockBatch
ElseIf SSTab1.Caption = "Create PO" Then
   SSTab1.Caption = "Create PO"
   RcPlan.DBOpen "SELECT [Planned Order].NoItem AS [Item ID], [Planned Order].[DESC] AS Description, [Planned Order].[Convert], " & _
                      "[Planned Order].M_OR_P, [Planned Order].PartnerID AS [Partner ID], PartnerDB.CompanyName AS Company, " & _
                      "[Planned Order].[Suggest QTY], [Planned Order].[Order QTY], [Planned Order].[Required Date], [Planned Order].[Suggest Order Date], " & _
                      "[Planned Order].[Order Date] , [Planned Order].ID, [Planned Order].OrderID, [Planned Order].Note " & _
                " FROM Inventory INNER JOIN [Planned Order] INNER JOIN " & _
                      "PartnerDB ON [Planned Order].PartnerID = PartnerDB.PartnerID ON Inventory.NoItem = [Planned Order].NoItem " & _
                " WHERE ([Planned Order].M_OR_P = 0) AND (Inventory.categid <> 'RL') " & _
                " ORDER BY [Partner ID], [Item ID]", CNN, lckLockBatch
Else
   SSTab1.Caption = "RPB"
   RcPlan.DBOpen "SELECT [Planned Order].NoItem AS [Item ID], [Planned Order].[DESC] AS Description, [Planned Order].[Convert], " & _
                      "[Planned Order].M_OR_P, [Planned Order].PartnerID AS [Partner ID], PartnerDB.CompanyName AS Company, " & _
                      "[Planned Order].[Suggest QTY], [Planned Order].[Order QTY], [Planned Order].[Required Date], [Planned Order].[Suggest Order Date], " & _
                      "[Planned Order].[Order Date] , [Planned Order].ID, [Planned Order].OrderID, [Planned Order].Note " & _
                " FROM Inventory INNER JOIN [Planned Order] INNER JOIN " & _
                      "PartnerDB ON [Planned Order].PartnerID = PartnerDB.PartnerID ON Inventory.NoItem = [Planned Order].NoItem " & _
                " WHERE ([Planned Order].M_OR_P = 0) AND (Inventory.categid = 'RL') " & _
                " ORDER BY [Partner ID], [Item ID]", CNN, lckLockBatch
End If

Set Dgdetail(SSTab1.Tab).DataSource = RcPlan.DBRecordset
Dgdetail(SSTab1.Tab).Columns(2).Button = True
Dgdetail(SSTab1.Tab).Columns(10).Button = True
End Sub

