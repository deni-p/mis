VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{82968C93-C596-4A47-8A14-646737648F29}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmMasterFixAssets 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Master Fixed Assets"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMasterFixAssets.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   Tag             =   "Master Fixed Asset"
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   15
      ScaleHeight     =   7185
      ScaleWidth      =   11115
      TabIndex        =   41
      Top             =   0
      Width           =   11145
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   6615
         Left            =   105
         ScaleHeight     =   6585
         ScaleWidth      =   10890
         TabIndex        =   42
         Top             =   105
         Width           =   10920
         Begin VB.TextBox txtBox 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "Aqui Cost"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0;(#.##0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Adodc1"
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
            Index           =   7
            Left            =   7365
            MaxLength       =   9
            TabIndex        =   16
            Tag             =   "Partner"
            Top             =   1545
            Width           =   3405
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Status Active"
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
            Height          =   270
            Left            =   4005
            TabIndex        =   2
            Top             =   225
            Width           =   1890
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "Date Added"
            Height          =   300
            Index           =   0
            Left            =   7365
            TabIndex        =   38
            Top             =   3885
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   529
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
            Format          =   57212931
            CurrentDate     =   38612
         End
         Begin VB.TextBox txtBox 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "Quantiy"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0;(#.##0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   1
            EndProperty
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
            Index           =   11
            Left            =   7365
            MaxLength       =   9
            TabIndex        =   30
            Tag             =   "Partner"
            Top             =   3225
            Width           =   3405
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "Man Name"
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
            Index           =   10
            Left            =   1965
            MaxLength       =   50
            TabIndex        =   36
            Tag             =   "Partner"
            Top             =   3885
            Width           =   3930
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "Custodian"
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
            Index           =   9
            Left            =   1965
            MaxLength       =   50
            TabIndex        =   32
            Tag             =   "Partner"
            Top             =   3555
            Width           =   3930
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "Alias Name"
            DataSource      =   "Adodc1"
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
            Index           =   6
            Left            =   1965
            MaxLength       =   50
            TabIndex        =   9
            Tag             =   "Partner"
            Top             =   1215
            Width           =   3930
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "Desc Ex"
            DataSource      =   "Adodc1"
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
            Index           =   5
            Left            =   7365
            MaxLength       =   100
            TabIndex        =   7
            Tag             =   "Partner"
            Top             =   885
            Width           =   3405
         End
         Begin VB.CommandButton cmdLink 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   5490
            Picture         =   "FrmMasterFixAssets.frx":6852
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   1545
            Width           =   405
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   2250
            Index           =   0
            Left            =   225
            TabIndex        =   39
            Tag             =   "Partner"
            Top             =   4275
            Width           =   10590
            _ExtentX        =   18680
            _ExtentY        =   3969
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   16777215
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
            ColumnCount     =   5
            BeginProperty Column00 
               DataField       =   "No Aktiva"
               Caption         =   "No Aktiva"
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
               DataField       =   "Nama Aktiva"
               Caption         =   "Nama Aktiva"
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
               DataField       =   "Lokasi"
               Caption         =   "Lokasi"
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
               DataField       =   "Departemen"
               Caption         =   "Departemen"
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
               DataField       =   "Serial"
               Caption         =   "Serial"
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
               EndProperty
            EndProperty
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "Serial"
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
            Left            =   1965
            MaxLength       =   50
            TabIndex        =   26
            Tag             =   "Partner"
            Top             =   2895
            Width           =   3930
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "Lokasi"
            DataSource      =   "Adodc1"
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
            Left            =   1965
            MaxLength       =   100
            TabIndex        =   18
            Tag             =   "Partner"
            Top             =   1875
            Width           =   3930
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "Departemen"
            DataSource      =   "Adodc1"
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
            Left            =   1965
            MaxLength       =   100
            TabIndex        =   13
            Tag             =   "Partner"
            Top             =   1545
            Width           =   3510
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "Nama Aktiva"
            DataSource      =   "Adodc1"
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
            Left            =   1965
            MaxLength       =   100
            TabIndex        =   6
            Tag             =   "Partner"
            Top             =   885
            Width           =   3930
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "No Aktiva"
            DataSource      =   "Adodc1"
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
            Left            =   1965
            MaxLength       =   16
            TabIndex        =   1
            Tag             =   "Partner"
            Top             =   210
            Width           =   1935
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            DataField       =   "NoAccount"
            Height          =   315
            Left            =   1965
            TabIndex        =   4
            Tag             =   "Partner"
            Top             =   540
            Width           =   3930
            _ExtentX        =   6932
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Nama Akun"
            BoundColumn     =   "NoAccount"
            Text            =   ""
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
         Begin MSDataListLib.DataCombo DataCombo1 
            DataField       =   "LocFisID"
            Height          =   315
            Index           =   1
            Left            =   1965
            TabIndex        =   22
            Tag             =   "Partner"
            Top             =   2565
            Width           =   3930
            _ExtentX        =   6932
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Description"
            BoundColumn     =   "LocFisID"
            Text            =   ""
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
         Begin MSDataListLib.DataCombo DataCombo1 
            DataField       =   "StrucID"
            Height          =   315
            Index           =   2
            Left            =   1965
            TabIndex        =   28
            Tag             =   "Partner"
            Top             =   3225
            Width           =   3930
            _ExtentX        =   6932
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Description"
            BoundColumn     =   "StrucID"
            Text            =   ""
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
         Begin MSDataListLib.DataCombo DataCombo1 
            DataField       =   "LocID"
            Height          =   315
            Index           =   3
            Left            =   7365
            TabIndex        =   24
            Tag             =   "Partner"
            Top             =   2580
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Description"
            BoundColumn     =   "LocID"
            Text            =   ""
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
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Index           =   1
            Left            =   7365
            TabIndex        =   11
            Top             =   1215
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   529
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
            Format          =   57212931
            CurrentDate     =   38612
         End
         Begin MSDataListLib.DataCombo CboUang 
            DataField       =   "CurrID"
            Height          =   330
            Left            =   7365
            TabIndex        =   20
            Tag             =   "PO"
            Top             =   1875
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Currency Name"
            BoundColumn     =   "CurrID"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "Last Main"
            Height          =   300
            Index           =   2
            Left            =   7365
            TabIndex        =   34
            Top             =   3555
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   529
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
            Format          =   57212931
            CurrentDate     =   38612
         End
         Begin VB.Line Line1 
            Index           =   11
            X1              =   5985
            X2              =   7785
            Y1              =   1185
            Y2              =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Extended Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   11
            Left            =   6015
            TabIndex        =   43
            Top             =   945
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aqusition  Cost"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   18
            Left            =   6015
            TabIndex        =   15
            Top             =   1605
            Width           =   1080
         End
         Begin VB.Line Line1 
            Index           =   18
            X1              =   5985
            X2              =   7785
            Y1              =   1845
            Y2              =   1845
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Currency ID"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   17
            Left            =   6015
            TabIndex        =   19
            Top             =   1935
            Width           =   870
         End
         Begin VB.Line Line1 
            Index           =   17
            X1              =   5985
            X2              =   7485
            Y1              =   2175
            Y2              =   2175
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aqusition Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   16
            Left            =   6015
            TabIndex        =   10
            Top             =   1260
            Width           =   1050
         End
         Begin VB.Line Line1 
            Index           =   16
            X1              =   5985
            X2              =   7785
            Y1              =   1500
            Y2              =   1500
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Location ID"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   8
            Left            =   6015
            TabIndex        =   23
            Top             =   2640
            Width           =   810
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   5985
            X2              =   7785
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   240
            X2              =   2040
            Y1              =   3525
            Y2              =   3525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Structure ID"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   27
            Top             =   3285
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phisycal Loc. ID"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   15
            Left            =   240
            TabIndex        =   21
            Top             =   2625
            Width           =   1125
         End
         Begin VB.Line Line1 
            Index           =   15
            X1              =   240
            X2              =   2040
            Y1              =   2865
            Y2              =   2865
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Added"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   14
            Left            =   6015
            TabIndex        =   37
            Top             =   3945
            Width           =   855
         End
         Begin VB.Line Line1 
            Index           =   14
            X1              =   5985
            X2              =   7785
            Y1              =   4170
            Y2              =   4170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last Maintenance"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   13
            Left            =   6015
            TabIndex        =   33
            Top             =   3615
            Width           =   1260
         End
         Begin VB.Line Line1 
            Index           =   13
            X1              =   5985
            X2              =   7785
            Y1              =   3840
            Y2              =   3840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   12
            Left            =   6015
            TabIndex        =   29
            Top             =   3285
            Width           =   630
         End
         Begin VB.Line Line1 
            Index           =   12
            X1              =   5985
            X2              =   7785
            Y1              =   3525
            Y2              =   3525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Manufacture Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   35
            Top             =   3945
            Width           =   1365
         End
         Begin VB.Line Line1 
            Index           =   10
            X1              =   240
            X2              =   2040
            Y1              =   4185
            Y2              =   4185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Custodian"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   31
            Top             =   3615
            Width           =   720
         End
         Begin VB.Line Line1 
            Index           =   9
            X1              =   240
            X2              =   2040
            Y1              =   3855
            Y2              =   3855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Alias"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   1
            Left            =   255
            TabIndex        =   8
            Top             =   1290
            Width           =   780
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   225
            X2              =   2025
            Y1              =   1515
            Y2              =   1515
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   240
            X2              =   2040
            Y1              =   3195
            Y2              =   3195
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   240
            X2              =   2040
            Y1              =   2175
            Y2              =   2175
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   240
            X2              =   2040
            Y1              =   1845
            Y2              =   1845
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   255
            X2              =   2055
            Y1              =   1185
            Y2              =   1185
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   240
            X2              =   2040
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   240
            X2              =   2040
            Y1              =   510
            Y2              =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Akun"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   6
            Left            =   270
            TabIndex        =   3
            Top             =   615
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Serial#"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   25
            Top             =   2955
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lokasi"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   4
            Left            =   270
            TabIndex        =   17
            Top             =   1935
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Departemen"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   3
            Left            =   270
            TabIndex        =   12
            Top             =   1605
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Aktiva "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   2
            Left            =   270
            TabIndex        =   5
            Top             =   960
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Aktiva"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   0
            Left            =   270
            TabIndex        =   0
            Top             =   255
            Width           =   690
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   40
      Top             =   7275
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmMasterFixAssets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RcGroup As New DBQuick
Private rcAkun As New DBQuick
Private RcPartner As New DBQuick
Private RcLocFis As New DBQuick
Private RcLoc As New DBQuick
Private RcCur As New DBQuick
Private RcStruc As New DBQuick
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private MyData As New clsTransaksi
Private Sub Form_Load()
HiasForm Picture1, Me
CenterForm Picture2, Me
GridLayout
Set mCall = New frmCaller
rcAkun.DBOpen "SELECT     NoAccount, AccountName AS [Nama Akun] FROM         GLAccount WHERE     ([Group] = N'Detail List Account') AND (Type = N'Aktiva Tetap Kantor' OR                       Type = N'Aktiva Tetap Produksi') ORDER BY NoAccount", CNN, lckLockReadOnly
Set DataCombo2.RowSource = rcAkun.DBRecordset

RcLocFis.DBOpen "SELECT     LocFisID, [Desc] AS Description FROM         SetupLocFisik", CNN, lckLockBatch
Set DataCombo1(1).RowSource = RcLocFis.DBRecordset

RcStruc.DBOpen "SELECT     StrucID, [Desc] AS Description FROM         SetupStructure", CNN, lckLockBatch
Set DataCombo1(2).RowSource = RcStruc.DBRecordset

RcLoc.DBOpen "SELECT     LocID, DescProp AS Description FROM         SetupLoc", CNN, lckLockBatch
Set DataCombo1(3).RowSource = RcLoc.DBRecordset

RcCur.DBOpen MyData.UploadQuery("mata uang"), CNN, lckLockBatch
Set CboUang.RowSource = RcCur.DBRecordset

With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmMasterFixAssets
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT  [Tabel Aktiva Tetap].[No Aktiva], [Tabel Aktiva Tetap].[Nama Aktiva], [Tabel Aktiva Tetap].Lokasi, [Tabel Aktiva Tetap].[Kode Dep] AS [Kode Departemen], [Tabel Departemen].[Nama Dep] AS Departemen, [Tabel Aktiva Tetap].Serial, [Tabel Aktiva Tetap].NoAccount FROM         [Tabel Aktiva Tetap] INNER JOIN                       [Tabel Departemen] ON [Tabel Aktiva Tetap].[Kode Dep] = [Tabel Departemen].[Kode Dep] ORDER BY [Tabel Aktiva Tetap].[No Aktiva]"
End With
End Sub
Private Sub cmdLink_Click(Index As Integer)
OpenPartner Index
End Sub

Private Sub DataCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub DataCombo2_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub DataGrid1_Error(Index As Integer, ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub

Private Sub DTPicker1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MyDDE.CheckRecordPendinged = True Then
   ScanKey vbKeyF5, 0, MyDDE
   If MyDDE.IsSucces = True Then
      Cancel = False
      MyDDE.ClearRecordset
      RcGroup.CloseDB
      rcAkun.CloseDB
      Set mCall = Nothing
   Else
      Cancel = True
   End If
Else
   MyDDE.ClearRecordset
   RcGroup.CloseDB
   rcAkun.CloseDB
   Set mCall = Nothing
End If
End Sub

Private Sub Form_Resize()

HiasForm Picture1, Me
CenterForm Picture2, Me
'GridLayout
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmMasterFixAssets = Nothing
End Sub

Private Sub mCall_BeforeUnload()
If txtBox(3).Enabled = True Then txtBox(3).SetFocus
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
MyDDE.GetFieldByName("Kode Departemen") = mCall.GetFieldByName(0)
MyDDE.GetFieldByName("Departemen") = mCall.GetFieldByName(1)
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbAddNew:
            txtBox(0).SetFocus
            txtBox(2).Enabled = False
            cmdLink(1).Enabled = True
       Case tmbEdit:
            txtBox(0).Enabled = False
            cmdLink(1).Enabled = True
            txtBox(2).Enabled = False
            DataCombo2.SetFocus
       Case tmbSave:
            cmdLink(1).Enabled = False
       Case tmbCancel:
            cmdLink(1).Enabled = False
       Case tmbDelete: cmdLink(1).Enabled = False
       Case tmbPrint:
            CallRPTReport "Daftar Aktiva.rpt"
       Case Else: 'mVarDataDc = False
End Select

End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim mDel As New clsDelete
txtBox(2).Enabled = False
Select Case AdReasonActiveDb
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               If mDel.CekDelete(txtBox(0), reDelMasterAktiva) = False Then
                  MyDDE.IsChildMemberReady = True
                  PrepareQuery
               Else
                  MyDDE.CancelTrans = True
                  MessageBox "Record (" & txtBox(0) & ") Sedang Dipakai Transaksi Lain." & vbCrLf & "Record Tidak Bisa DiHapus.", "Peringatan", msgOkOnly
                  MyDDE.IsChildMemberReady = False
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               MyDDE.GetFieldByName("NoAccount") = DataCombo2.BoundText
               PrepareQuery
               
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
'Set mDel = Nothing
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Picture1.Parent.hwnd
End Sub

Private Sub txtBox_Change(Index As Integer)
Select Case Index
       Case 7, 11: If txtBox(Index) = "" Then txtBox(Index) = 0
       Case Else: If txtBox(Index) = "" Then txtBox(Index) = "-"
End Select

End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub PrepareQuery()
With MyDDE
    .PrepareAppend = " INSERT INTO [Tabel Aktiva Tetap] ( [No Aktiva],  [Nama Aktiva], Lokasi, [Kode Dep], Serial,NoAccount,[Desc Ex], [Alias Name], [Aqui Cost], CurrID, LocID, LocFisID, StrucID, Custodian, [Man Name], Quantiy, [Last Main], [Date Added],[Aqui Date],[Status]) " & _
                     " VALUES (N'" & ValidString(txtBox(0)) & "',N'" & ValidString(txtBox(1)) & "',N'" & ValidString(txtBox(3)) & "'," & MyDDE.GetFieldByName("Kode Departemen") & ",N'" & ValidString(txtBox(4)) & "',N'" & DataCombo2.BoundText & "',N'" & ValidString(txtBox(5)) & "',N'" & ValidString(txtBox(6)) & "'," & CDbl(txtBox(7)) & ",N'" & CboUang.BoundText & "'," & _
                     " N'" & DataCombo1(3).BoundText & "', N'" & DataCombo1(1).BoundText & "',N'" & DataCombo1(2).BoundText & "',N'" & ValidString(txtBox(9)) & "',N'" & ValidString(txtBox(10)) & "'," & CDbl(txtBox(11)) & ",Convert(Datetime,'" & Format(DTPicker1(2).Value, "dd/mm/yy") & "',3),Convert(Datetime,'" & Format(DTPicker1(0).Value, "dd/mm/yy") & "',3),Convert(Datetime,'" & Format(DTPicker1(1).Value, "dd/mm/yy") & "',3)," & BoolToInt(CBool(Check1.Value)) & ")"
                     
    .PrepareUpdate = " UPDATE [Tabel Aktiva Tetap] Set [Nama Aktiva] = N'" & ValidString(txtBox(1)) & "', [Kode Dep] = " & MyDDE.GetFieldByName("Kode Departemen") & ", Lokasi = N'" & ValidString(txtBox(3)) & "',NoAccount = N'" & DataCombo2.BoundText & "', Serial = N'" & ValidString(txtBox(4)) & "',[Status] =" & BoolToInt(CBool(Check1.Value)) & ", " & _
                     " [Desc Ex] = N'" & ValidString(txtBox(5)) & "', [Alias Name]= N'" & ValidString(txtBox(6)) & "', [Aqui Cost] = " & CDbl(txtBox(7)) & ", CurrID = N'" & CboUang.BoundText & "'," & _
                     " LocID=N'" & DataCombo1(3).BoundText & "', LocFisID = N'" & DataCombo1(1).BoundText & "',StrucID = N'" & DataCombo1(2).BoundText & "', Custodian = N'" & ValidString(txtBox(9)) & "',[Man Name]=N'" & ValidString(txtBox(10)) & "', Quantiy = " & CDbl(txtBox(11)) & ", [Last Main] = Convert(Datetime,'" & Format(DTPicker1(2).Value, "dd/mm/yy") & "',3),[Date Added]=Convert(Datetime,'" & Format(DTPicker1(0).Value, "dd/mm/yy") & "',3),[Aqui Date] = Convert(Datetime,'" & Format(DTPicker1(1).Value, "dd/mm/yy") & "',3)" & _
                     " WHERE  ([No Aktiva] = N'" & ValidString(txtBox(0)) & "')"

    .PrepareDelete = " DELETE FROM [Tabel Aktiva Tetap] WHERE   ([No Aktiva] = N'" & ValidString(txtBox(0)) & "') "
End With
End Sub

Private Sub GridLayout()
DataGrid1(0).Height = 2250
DataGrid1(0).Width = 10590
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:
Select Case Index
       Case 1:
            RcPartner.DBOpen " SELECT     [Kode Dep] AS [Kode Departemen], [Nama Dep] AS Departemen FROM         [Tabel Departemen] WHERE     (Type = 0) AND (ReportsTo <> N'TOP') ORDER BY [Nama Dep]", CNN, lckLockReadOnly
'       Case 2:
'            RcPartner.DBOpen "SELECT Inventory.NoItem, Inventory.ItemName, Inventory.UOM, Inventory.PPn, MAX([Inventory Tabel].PriceIn) * (Inventory.PPn / 100)  + MAX([Inventory Tabel].PriceIn) * (Inventory.Markup / 100) + MAX([Inventory Tabel].PriceIn) AS Harga, SUM([Inventory Tabel].QTY_IN) AS QTY FROM Inventory LEFT OUTER JOIN [Inventory Tabel] ON Inventory.NoItem = [Inventory Tabel].NoItem WHERE     ([Inventory Tabel].LockFIFO = 0) GROUP BY Inventory.NoItem, Inventory.ItemName, Inventory.PPn, Inventory.Markup, Inventory.UOM HAVING      (SUM([Inventory Tabel].QTY_IN) <> 0)", Cnn, lckLockReadOnly
'            DGPurchase.Columns(6).Visible = False
'            DGPurchase.Columns(7).Visible = True
            
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 1:
            mCall.FromTagActive = "TABEL DEPARTMENT"
'            mCall.txtCari = txtBox(3)
'          Case 2:
'            mCall.FromTagActive = "MASTER BARANG"
'            mCall.txtCari = txtBox(2)
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me
Else
   MessageBox "Data masih kosong.", "Peringatan", msgOkOnly
End If
Exit Sub
Hell:
    Err.Clear
End Sub

Private Sub txtBox_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
       Case 7, 11: ValidNum KeyAscii
       Case Else:
End Select
End Sub
