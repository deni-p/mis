VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmInvMemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Invoice Memo"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11040
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvMemo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   Tag             =   "Credit / Debit Memo"
   Begin VB.PictureBox Picture2 
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
      Height          =   5955
      Left            =   0
      ScaleHeight     =   5955
      ScaleWidth      =   11040
      TabIndex        =   11
      Top             =   0
      Width           =   11040
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Keterangan"
         Height          =   330
         Left            =   1440
         MaxLength       =   200
         TabIndex        =   9
         Tag             =   "ASM"
         Top             =   5415
         Width           =   3870
      End
      Begin VB.CommandButton cmdLInk 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4305
         Picture         =   "frmInvMemo.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   803
         Width           =   345
      End
      Begin MSDataGridLib.DataGrid DgDetail 
         Bindings        =   "frmInvMemo.frx":6BDC
         Height          =   1665
         Left            =   120
         TabIndex        =   8
         Top             =   3705
         Width           =   10620
         _ExtentX        =   18733
         _ExtentY        =   2937
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16777215
         HeadLines       =   2
         RowHeight       =   15
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
         BeginProperty Column03 
            DataField       =   "QTY Beli"
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
         BeginProperty Column04 
            DataField       =   "PPN"
            Caption         =   "PPN"
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
         BeginProperty Column05 
            DataField       =   "Harga"
            Caption         =   "Harga"
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
         BeginProperty Column06 
            DataField       =   "Total"
            Caption         =   "Total"
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
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DGPurchase 
         Bindings        =   "frmInvMemo.frx":6BF1
         Height          =   1665
         Left            =   105
         TabIndex        =   6
         Top             =   1620
         Width           =   10620
         _ExtentX        =   18733
         _ExtentY        =   2937
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16777215
         HeadLines       =   2
         RowHeight       =   15
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
         BeginProperty Column03 
            DataField       =   "QTY Beli"
            Caption         =   "QTY Beli"
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
         BeginProperty Column04 
            DataField       =   "PPN"
            Caption         =   "PPN"
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
         BeginProperty Column05 
            DataField       =   "Harga"
            Caption         =   "Harga"
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
         BeginProperty Column06 
            DataField       =   "Total"
            Caption         =   "Total"
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
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "Tanggal bukti"
         Height          =   330
         Left            =   1215
         TabIndex        =   2
         Tag             =   "ASM"
         Top             =   420
         Width           =   3090
         _ExtentX        =   5450
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
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
         Left            =   105
         TabIndex        =   28
         Top             =   5475
         Width           =   1065
      End
      Begin VB.Label lblFixAssets 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No Bukti"
         DataField       =   "No Journal"
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
         Left            =   1215
         TabIndex        =   1
         Tag             =   "ASM"
         Top             =   60
         Width           =   3090
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
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
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   27
         Top             =   488
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Bukti"
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
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   26
         Top             =   128
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Left            =   7650
         TabIndex        =   25
         Top             =   4620
         Width           =   465
      End
      Begin VB.Label LblAmount 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   10605
         TabIndex        =   24
         Top             =   4665
         Width           =   120
      End
      Begin VB.Label lblFixAssets 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No Bukti"
         DataField       =   "No Bukti"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   1215
         TabIndex        =   3
         Tag             =   "ASM"
         Top             =   795
         Width           =   3090
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref Bukti"
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
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   23
         Top             =   868
         Width           =   765
      End
      Begin VB.Label lblFixAssets 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No Bukti"
         DataField       =   "Doc Reff"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   2
         Left            =   1215
         TabIndex        =   5
         Tag             =   "ASM"
         Top             =   1170
         Width           =   3090
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Doc Reff"
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
         Height          =   195
         Index           =   5
         Left            =   105
         TabIndex        =   22
         Top             =   1238
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grand Total"
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
         Index           =   6
         Left            =   6465
         TabIndex        =   21
         Top             =   3375
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grand Total"
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
         Left            =   6480
         TabIndex        =   20
         Top             =   5475
         Width           =   1065
      End
      Begin VB.Label LblTotalSource 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   7920
         TabIndex        =   7
         Top             =   3315
         Width           =   2790
      End
      Begin VB.Label LblTotalSource 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   7920
         TabIndex        =   10
         Top             =   5415
         Width           =   2790
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Partner"
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
         Index           =   8
         Left            =   5550
         TabIndex        =   19
         Top             =   165
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Perusahaan"
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
         Left            =   5550
         TabIndex        =   18
         Top             =   480
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
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
         Index           =   10
         Left            =   5550
         TabIndex        =   17
         Top             =   780
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Term"
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
         Index           =   11
         Left            =   5550
         TabIndex        =   16
         Top             =   1095
         Width           =   450
      End
      Begin VB.Label lblFixAssets 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Bukti"
         DataField       =   "Kode Partner"
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
         Left            =   7290
         TabIndex        =   15
         Tag             =   "ASM"
         Top             =   150
         Width           =   780
      End
      Begin VB.Label lblFixAssets 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Bukti"
         DataField       =   "Nama Perusahaan"
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
         Left            =   7290
         TabIndex        =   14
         Tag             =   "ASM"
         Top             =   480
         Width           =   780
      End
      Begin VB.Label lblFixAssets 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Bukti"
         DataField       =   "Alamat"
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
         Left            =   7290
         TabIndex        =   13
         Tag             =   "ASM"
         Top             =   795
         Width           =   780
      End
      Begin VB.Label lblFixAssets 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Bukti"
         DataField       =   "Term"
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
         Index           =   6
         Left            =   7290
         TabIndex        =   12
         Tag             =   "ASM"
         Top             =   1110
         Width           =   780
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   105
         X2              =   1725
         Y1              =   375
         Y2              =   375
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   105
         X2              =   1725
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   105
         X2              =   1725
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   105
         X2              =   1725
         Y1              =   1485
         Y2              =   1485
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   105
         X2              =   1725
         Y1              =   5730
         Y2              =   5730
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   6465
         X2              =   8085
         Y1              =   5730
         Y2              =   5730
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   6450
         X2              =   8070
         Y1              =   3630
         Y2              =   3630
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5955
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmInvMemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RcDetail As New DBQuick
Private RcGroup As New DBQuick
Private MyData As New clsTransaksi
Private mVarAdd As Boolean
Private mVarTmp As String
Private mVarJournal As String
Private mVarKas As Variant
Private mVarPPn As Variant
Private mVarBarang As Variant
Private mVarHpp As Variant
Private mVarHarga As Variant

Private Sub cmdLink_Click()
OpenPartner 0
End Sub

Private Sub DGDetail_AfterColEdit(ByVal ColIndex As Integer)
Dim mTotalTrans As Currency
Dim mPPn As Single
If mVarAdd = True Then
   Select Case dgDetail.col
          Case 3, 4, 5:
               With MyDDE.ChildRecordset
               If .Fields("ppn") <> 0 Then
                  mPPn = .Fields("PPn") / 100
               Else
                  mPPn = 0
               End If
               mTotalTrans = ((dgDetail.Columns(5).Value * mPPn) + dgDetail.Columns(5).Value) * dgDetail.Columns(3).Value
               dgDetail.Columns(6).Value = mTotalTrans
               End With
               TotalKas
          Case Else:
   End Select
End If
End Sub

Private Sub DGDETAIL_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If mVarAdd = True Then
   Select Case dgDetail.col
          Case 3, 4, 5: dgDetail.AllowUpdate = True
          Case Else: dgDetail.AllowUpdate = False
   End Select
Else
   dgDetail.AllowUpdate = False
End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Form_Activate()
'If Me.WindowState = 0 Then Me.WindowState = 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
End Sub

Private Sub Form_Load()
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
DTPicker1.Value = dDateBegin
With MyDDE
    .EditModeReplace = False
    Set .BindForm = frmInvMemo
    .BindFormTAG = "ASM"
    Set .ActiveConnection = CNN
    .PrepareQuery = "select [Table Journal].JournalID AS [No Journal], [Table Journal].TransID AS [No Bukti], [Table Journal].PurchaseID AS [Doc Reff],                        [Table Journal].DateTrans AS [Tanggal Bukti], [Table Journal].RefNotes AS Keterangan, PartnerDB.PartnerID AS [Kode Partner],                        PartnerDB.CompanyName AS [Nama Perusahaan], PartnerDB.Address AS Alamat, [PO Order].TermPayment AS Term FROM         [Table Journal] INNER JOIN                       [PO Order] ON [Table Journal].PurchaseID = [PO Order].PurchaseID INNER JOIN                       PartnerDB ON [PO Order].PartnerID = PartnerDB.PartnerID WHERE     ([Table Journal].TypeTrans = N'INVMEMO') AND ([Table Journal].Periode = " & mVarPeriode & ") ORDER BY [Table Journal].JournalID"
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
RcGroup.CloseDB
MyDDE.ClearRecordset
Set MyData = Nothing
Set mCall = Nothing
End Sub

Private Sub Form_Resize()

'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmInvMemo = Nothing
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case UCase(TagForm)
       Case "DOKUMEN TRANSAKSI":
            With MyDDE
                 mVarJournal = mCall.GetFieldByName(1)
                 .GetFieldByName("No Bukti") = mCall.GetFieldByName(2)
                 .GetFieldByName("Doc Reff") = mCall.GetFieldByName(3)
                 .GetFieldByName("Keterangan") = "Koreksi Data - " & .GetFieldByName("No Bukti")
                 OpenDetailLama
                 TotalKas
            End With
       Case "DETAIL TRANSAKSI":
            With MyDDE.ChildRecordset
                 .Fields(0) = mCall.GetFieldByName(0)
                 .Fields(1) = mCall.GetFieldByName(1)
                 .Fields(2) = mCall.GetFieldByName(2)
                 .Fields(3) = 0
                 .Fields(4) = mCall.GetFieldByName(4)
                 .Fields(5) = mCall.GetFieldByName(5)
            End With
            
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit:
       Case tmbAddNew:
            DTPicker1.Value = CDate(Format(dDateBegin, "dd/mm/yyyy"))
            With MyDDE
                 .GetFieldByName("No Journal") = MyData.PrepareIndex(tmbTransaksiinvMemorial, 5, "", TglIndex)
                 .GetFieldByName("Keterangan") = "Koreksi Data -" & lblFixAssets(1)
                 .GetFieldByName("Tanggal bukti") = DTPicker1.Value
            End With
            mVarAdd = True
            OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("No Bukti")), MyDDE.GetFieldByName("No Bukti"), "XXXXX"), IIf(Not IsNull(MyDDE.GetFieldByName("Doc Reff")), MyDDE.GetFieldByName("Doc Reff"), "XXXXX")
       Case tmbDetail:
            mVarAdd = True
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               If OpenPartner(1) = True Then CancelDetailTrans
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbDelete:
            If MyDDE.IsChildMemberReady = True Then
               'SendDataToServer ("DELETE FROM Inventory WHERE     (NoItem = N'" & 'txtBox(0) & "') ")
            End If
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               SimpanDetail
            End If
       Case tmbPrint:
            CallRPTReport "Bukti Memorial.Rpt", "Select * from [bukti memorial] Where [No Bukti]='" & lblFixAssets(0) & "'"
End Select
cmdLink.Enabled = mVarAdd
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("No Bukti")), MyDDE.GetFieldByName("No Bukti"), "XXXXX"), IIf(Not IsNull(MyDDE.GetFieldByName("Doc Reff")), MyDDE.GetFieldByName("Doc Reff"), "XXXXX")
TotalKas
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
                If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.CancelTrans = False
                   If MyDDE.CancelTrans = True Then
                      MyDDE.IsChildMemberReady = False
                      MessageBox "Data detail belum Balance. Harap diperiksa dulu.", "Peringatan", msgOkOnly, msgCrtical
                   Else
                      MyDDE.IsChildMemberReady = True
                      'SimpanDetail
                      mVarAdd = False
                   End If
            Else
               MessageBox "Data detail belum Lengkap. Harap diisi dulu.", "Peringatan", msgOkOnly, msgCrtical
            End If
'            If MyDDE.CheckEmptyControl = False Then
'               If MyDDE.ChildRecordset.Recordcount <> 0 Then
'                  MyDDE.IsChildMemberReady = True
'               Else
'                  MyDDE.IsChildMemberReady = False
'                  MessageBox "Data detail belum ada. Harap diisi dulu.", "Peringatan", msgOkOnly
'               End If
'            Else
'               MyDDE.IsChildMemberReady = False
'            End If
            
       
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
            Else
               MyDDE.IsChildMemberReady = False
            End If
            mVarAdd = False
       Case tmbCancel:
            mVarAdd = False
'       Case tmbDetail:
'            If MyDDE.CheckEmptyControl = False Then
'               MyDDE.IsChildMemberReady = True
'               OpenPartner
'            Else
'               MyDDE.IsChildMemberReady = False
'            End If
       Case tmbSave:

End Select
End Sub

Private Sub PrepareQuery()
On Error GoTo xErr
With MyDDE
    .PrepareAppend = " INSERT INTO [Table Journal]" & _
                     " (JournalID, TransID, PurchaseID, DateTrans, Periode, TypeTrans, RefNotes)" & _
                     " VALUES (N'" & lblFixAssets(0) & "', N'" & lblFixAssets(1) & "', N'" & lblFixAssets(2) & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), " & mVarPeriode & ", N'INVMEMO', N'" & ValidString(Text1) & "')"

    .PrepareUpdate = " UPDATE    [Table Journal]" & _
                     " SET DateTrans = CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), Periode = " & mVarPeriode & ", TypeTrans = N'INVMEMO', RefNotes = N'" & ValidString(Text1) & "'" & _
                     " WHERE     (JournalID = N'" & lblFixAssets(0) & "') AND (TransID = N'" & lblFixAssets(1) & "') AND (PurchaseID = N'" & lblFixAssets(2) & "')"

    .PrepareDelete = " DELETE FROM [Table Journal] WHERE     (JournalID = N'" & lblFixAssets(0) & "') "
End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub SimpanDetail()
Dim RcJournal As New DBQuick
Dim mVarDR As Variant
Dim mVarCR As Variant
Dim mTotalBarang As Variant
On Error GoTo xErr
RcJournal.DBOpen " SELECT  [Detail Journal].NoAccount, [Detail Journal].[Doc Reff], [Detail Journal].Debet, [Detail Journal].Credit, [Detail Journal].Keterangan" & _
                 " FROM  [Detail Journal] INNER JOIN" & _
                 " [Table Journal] ON [Detail Journal].JournalID = [Table Journal].JournalID" & _
                 " WHERE ([Detail Journal].JournalID = N'" & mVarJournal & "') AND ([Table Journal].TransID = N'" & MyDDE.GetFieldByName("No Bukti") & "')", CNN, lckLockReadOnly
                 
With RcJournal.DBRecordset
     If .Recordcount <> 0 Then
        'Copy Journal
        Do
        If .EOF Then Exit Do
            SendDataToServer (" INSERT INTO [Detail Journal] " & _
                              " (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan)" & _
                              " VALUES (N'" & lblFixAssets(0) & "', N'" & .Fields("NoAccount") & "', N'" & .Fields("Doc Reff") & "', " & CCur(.Fields("Credit")) & ", " & .Fields("Debet") & ", N'" & .Fields("Keterangan") & "')")
           .MoveNext
        Loop
        .MoveFirst

        Do
        'Journal Balik Sebagai Pengenol
        If .EOF Then Exit Do
            SendDataToServer (" INSERT INTO [Detail Journal] " & _
                              " (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan)" & _
                              " VALUES (N'" & lblFixAssets(0) & "', N'" & .Fields("NoAccount") & "', N'" & .Fields("Doc Reff") & "', " & CCur(.Fields("Debet")) & ", " & .Fields("Credit") & ", N'" & .Fields("Keterangan") & "')")
           .MoveNext
        Loop
        .MoveFirst
        
     End If
End With

With MyDDE.ChildRecordset
     'Journal Pembelian
     If Left(lblFixAssets(1), 2) = "RN" Then
        TotalTransJournal
        If Val(lblFixAssets(6)) = 0 Then
           RcJournal.DBOpen "SELECT     NoAccount, Posisi, [Nama Form], [Value Data] FROM         [Daftar Configurasi] WHERE     ([Kode Konfigurasi] = N'BKKPTP') ORDER BY [No Index]", CNN, lckLockReadOnly
           If RcJournal.Recordcount <> 0 Then
              Do
                If RcJournal.DBRecordset.EOF = True Then Exit Do
                   Select Case UCase(RcJournal.DBRecordset.Fields("NAMA FORM"))
                          Case "KAS KELUAR":
                                mVarDR = mVarKas
                                mVarCR = mVarDR
                                If RcJournal.DBRecordset.Fields("Posisi") = False Then mVarDR = 0 Else mVarCR = 0
                                SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & lblFixAssets(0) & "', N'" & RcJournal.DBRecordset.Fields("NoAccount") & "', N'xxx', " & CCur(mVarDR) & " , " & CCur(mVarCR) & ", N'" & ValidString(MyDDE.GetFieldByName("Keterangan")) & "')")
                          Case "PPN MASUKAN":
                                mVarDR = mVarPPn
                                mVarCR = mVarDR
                                If RcJournal.DBRecordset.Fields("Posisi") = False Then mVarDR = 0 Else mVarCR = 0
                                SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & lblFixAssets(0) & "', N'" & RcJournal.DBRecordset.Fields("NoAccount") & "', N'xxx', " & CCur(mVarDR) & ", " & CCur(mVarCR) & ", N'" & ValidString(MyDDE.GetFieldByName("Keterangan")) & "')")
                          Case "TOTAL PEMBELIAN":
                               .MoveFirst
                               Do
                                 If .EOF = True Then Exit Do
                                 mVarBarang = .Fields(3) * .Fields(5)
                                 mVarDR = mVarBarang
                                 mVarCR = mVarDR
                                 If RcJournal.DBRecordset.Fields("Posisi") = False Then mVarDR = 0 Else mVarCR = 0
                                 SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & lblFixAssets(0) & "', N'" & RcJournal.DBRecordset.Fields("NoAccount") & "', N'" & .Fields(0) & "', " & CCur(mVarDR) & ", " & CCur(mVarCR) & ", N'" & ValidString(MyDDE.GetFieldByName("Keterangan")) & "')")
                                 mTotalBarang = TotalBarang
                                 SendARItem .Fields(0), CDbl(mTotalBarang), CDbl(mVarHarga), lblFixAssets(1), DTPicker1.Value, CCur(mVarHarga), "INVMEMO"
                                 'SendAPItem .Fields(0), CDbl(MyDDE.ChildRecordset.Fields(3)), CDbl(MyDDE.ChildRecordset.Fields(5)), lblFixAssets(0), DTPicker1.Value, "INVMEMO"
                                 .MoveNext
                               Loop
                               .MoveLast
                   End Select
                   RcJournal.DBRecordset.MoveNext
                Loop
           End If
        Else
           RcJournal.DBOpen "SELECT     NoAccount, Posisi, [Nama Form], [Value Data] FROM         [Daftar Configurasi] WHERE     ([Kode Konfigurasi] = N'BPBK') ORDER BY [No Index]", CNN, lckLockReadOnly
           If RcJournal.Recordcount <> 0 Then
              Do
                If RcJournal.DBRecordset.EOF = True Then Exit Do
                   Select Case UCase(RcJournal.DBRecordset.Fields("NAMA FORM"))
                          Case "HUTANG USAHA":
                                mVarDR = mVarKas
                                mVarCR = mVarDR
                                If RcJournal.DBRecordset.Fields("Posisi") = False Then mVarDR = 0 Else mVarCR = 0
                                SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & lblFixAssets(0) & "', N'" & RcJournal.DBRecordset.Fields("NoAccount") & "', N'" & lblFixAssets(3) & "', " & CCur(mVarDR) & " , " & CCur(mVarCR) & ", N'" & ValidString(MyDDE.GetFieldByName("Keterangan")) & "')")
                          Case "PPN MASUKAN":
                                mVarDR = mVarPPn
                                mVarCR = mVarDR
                                If RcJournal.DBRecordset.Fields("Posisi") = False Then mVarDR = 0 Else mVarCR = 0
                                SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & lblFixAssets(0) & "', N'" & RcJournal.DBRecordset.Fields("NoAccount") & "', N'xxx', " & CCur(mVarDR) & ", " & CCur(mVarCR) & ", N'" & ValidString(MyDDE.GetFieldByName("Keterangan")) & "')")
                          Case "TOTAL PEMBELIAN":
                               .MoveFirst
                               Do
                                 If .EOF = True Then Exit Do
                                 mVarBarang = .Fields(3) * .Fields(5)
                                 mVarDR = mVarBarang
                                 mVarCR = mVarDR
                                 If RcJournal.DBRecordset.Fields("Posisi") = False Then mVarDR = 0 Else mVarCR = 0
                                 SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & lblFixAssets(0) & "', N'" & RcJournal.DBRecordset.Fields("NoAccount") & "', N'" & .Fields(0) & "', " & CCur(mVarDR) & ", " & CCur(mVarCR) & ", N'" & ValidString(MyDDE.GetFieldByName("Keterangan")) & "')")
                                 SendARItem .Fields(0), CDbl(mTotalBarang), CDbl(mVarHarga), lblFixAssets(1), DTPicker1.Value, CCur(mVarHarga), "INVMEMO"
                                 'SendAPItem .Fields(0), CDbl(MyDDE.ChildRecordset.Fields(3)), CDbl(MyDDE.ChildRecordset.Fields(5)), lblFixAssets(0), DTPicker1.Value, "INVMEMO"
                                 .MoveNext
                               Loop
                               .MoveLast
                   End Select
                   RcJournal.DBRecordset.MoveNext
                Loop
           End If
        End If
     Else
        'Journal Penjualan
        TotalTransJournal
        If Val(lblFixAssets(6)) = 0 Then
           RcJournal.DBOpen "SELECT     NoAccount, Posisi, [Nama Form], [Value Data] FROM         [Daftar Configurasi] WHERE     ([Kode Konfigurasi] = N'BKMPTP') ORDER BY [No Index]", CNN, lckLockReadOnly
           If RcJournal.Recordcount <> 0 Then
              Do
                If RcJournal.DBRecordset.EOF = True Then Exit Do
                   Select Case UCase(RcJournal.DBRecordset.Fields("NAMA FORM"))
                          Case "KAS MASUK":
                                mVarDR = mVarKas
                                mVarCR = mVarDR
                                If RcJournal.DBRecordset.Fields("Posisi") = False Then mVarDR = 0 Else mVarCR = 0
                                SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & lblFixAssets(0) & "', N'" & RcJournal.DBRecordset.Fields("NoAccount") & "', N'xxx', " & CCur(mVarDR) & " , " & CCur(mVarCR) & ", N'" & ValidString(MyDDE.GetFieldByName("Keterangan")) & "')")
                                
                          Case "PPN KELUARAN":
                                mVarDR = mVarPPn
                                mVarCR = mVarDR
                                If RcJournal.DBRecordset.Fields("Posisi") = False Then mVarDR = 0 Else mVarCR = 0
                                SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & lblFixAssets(0) & "', N'" & RcJournal.DBRecordset.Fields("NoAccount") & "', N'xxx', " & CCur(mVarDR) & ", " & CCur(mVarCR) & ", N'" & ValidString(MyDDE.GetFieldByName("Keterangan")) & "')")
                                
                          Case "HPP":
                                mVarDR = mVarHpp
                                mVarCR = mVarDR
                                If RcJournal.DBRecordset.Fields("Posisi") = False Then mVarDR = 0 Else mVarCR = 0
                                SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & lblFixAssets(0) & "', N'" & RcJournal.DBRecordset.Fields("NoAccount") & "', N'xxx', " & CCur(mVarDR) & ", " & CCur(mVarCR) & ", N'" & ValidString(MyDDE.GetFieldByName("Keterangan")) & "')")
                                
                          Case "PENGHASILAN TUNAI":
                                mVarDR = mVarHpp
                                mVarCR = mVarDR
                                If RcJournal.DBRecordset.Fields("Posisi") = False Then mVarDR = 0 Else mVarCR = 0
                                SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & lblFixAssets(0) & "', N'" & RcJournal.DBRecordset.Fields("NoAccount") & "', N'xxx', " & CCur(mVarDR) & ", " & CCur(mVarCR) & ", N'" & ValidString(MyDDE.GetFieldByName("Keterangan")) & "')")
                                
                          Case "TOTAL PENJUALAN":
                               .MoveFirst
                               Do
                                 If .EOF = True Then Exit Do
                                 mVarBarang = .Fields(3) * .Fields(5)
                                 mVarDR = mVarBarang
                                 mVarCR = mVarDR
                                 If RcJournal.DBRecordset.Fields("Posisi") = False Then mVarDR = 0 Else mVarCR = 0
                                 SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & lblFixAssets(0) & "', N'" & RcJournal.DBRecordset.Fields("NoAccount") & "', N'" & .Fields(0) & "', " & CCur(mVarDR) & ", " & CCur(mVarCR) & ", N'" & ValidString(MyDDE.GetFieldByName("Keterangan")) & "')")
                                 mTotalBarang = TotalBarang
                                 'SendAPItem .Fields(0), CDbl(mTotalBarang), CDbl(mVarHarga), lblFixAssets(1), DTPicker1.Value, "INVMEMO"
                                 SendARItem .Fields(0), CDbl(MyDDE.ChildRecordset.Fields(3)), CDbl(MyDDE.ChildRecordset.Fields(5)), lblFixAssets(1), DTPicker1.Value, CCur(mVarHarga), "INVMEMO"
                                 
                                 .MoveNext
                               Loop
                               .MoveLast
                   End Select
                   RcJournal.DBRecordset.MoveNext
                Loop
           End If
        Else
           RcJournal.DBOpen "SELECT     NoAccount, Posisi, [Nama Form], [Value Data] FROM         [Daftar Configurasi] WHERE     ([Kode Konfigurasi] = N'BPBK') ORDER BY [No Index]", CNN, lckLockReadOnly
           If RcJournal.Recordcount <> 0 Then
              Do
                If RcJournal.DBRecordset.EOF = True Then Exit Do
                   Select Case UCase(RcJournal.DBRecordset.Fields("NAMA FORM"))
                          Case "KAS MASUK":
                                mVarDR = mVarKas
                                mVarCR = mVarDR
                                If RcJournal.DBRecordset.Fields("Posisi") = False Then mVarDR = 0 Else mVarCR = 0
                                SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & lblFixAssets(0) & "', N'" & RcJournal.DBRecordset.Fields("NoAccount") & "', N'xxx', " & CCur(mVarDR) & " , " & CCur(mVarCR) & ", N'" & ValidString(MyDDE.GetFieldByName("Keterangan")) & "')")
                                
                          Case "PPN KELUARAN":
                                mVarDR = mVarPPn
                                mVarCR = mVarDR
                                If RcJournal.DBRecordset.Fields("Posisi") = False Then mVarDR = 0 Else mVarCR = 0
                                SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & lblFixAssets(0) & "', N'" & RcJournal.DBRecordset.Fields("NoAccount") & "', N'xxx', " & CCur(mVarDR) & ", " & CCur(mVarCR) & ", N'" & ValidString(MyDDE.GetFieldByName("Keterangan")) & "')")
                                
                          Case "HPP":
                                mVarDR = mVarHpp
                                mVarCR = mVarDR
                                If RcJournal.DBRecordset.Fields("Posisi") = False Then mVarDR = 0 Else mVarCR = 0
                                SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & lblFixAssets(0) & "', N'" & RcJournal.DBRecordset.Fields("NoAccount") & "', N'xxx', " & CCur(mVarDR) & ", " & CCur(mVarCR) & ", N'" & ValidString(MyDDE.GetFieldByName("Keterangan")) & "')")
                                
                          Case "PENGHASILAN TUNAI":
                                mVarDR = mVarHpp
                                mVarCR = mVarDR
                                If RcJournal.DBRecordset.Fields("Posisi") = False Then mVarDR = 0 Else mVarCR = 0
                                SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & lblFixAssets(0) & "', N'" & RcJournal.DBRecordset.Fields("NoAccount") & "', N'xxx', " & CCur(mVarDR) & ", " & CCur(mVarCR) & ", N'" & ValidString(MyDDE.GetFieldByName("Keterangan")) & "')")
                                
                          Case "TOTAL PENJUALAN":
                               .MoveFirst
                               Do
                                 If .EOF = True Then Exit Do
                                 mVarBarang = .Fields(3) * .Fields(5)
                                 mVarDR = mVarBarang
                                 mVarCR = mVarDR
                                 If RcJournal.DBRecordset.Fields("Posisi") = False Then mVarDR = 0 Else mVarCR = 0
                                 SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & lblFixAssets(0) & "', N'" & RcJournal.DBRecordset.Fields("NoAccount") & "', N'" & .Fields(0) & "', " & CCur(mVarDR) & ", " & CCur(mVarCR) & ", N'" & ValidString(MyDDE.GetFieldByName("Keterangan")) & "')")
                                 mTotalBarang = TotalBarang
                                 'SendAPItem .Fields(0), CDbl(mTotalBarang), CDbl(mVarHarga), lblFixAssets(0), DTPicker1.Value, "INVMEMO"
                                 SendARItem .Fields(0), CDbl(MyDDE.ChildRecordset.Fields(3)), CDbl(MyDDE.ChildRecordset.Fields(5)), lblFixAssets(1), DTPicker1.Value, CCur(mVarHarga), "INVMEMO"
                                 .MoveNext
                               Loop
                               .MoveLast
                   End Select
                   RcJournal.DBRecordset.MoveNext
                Loop
           End If
        End If
     
     End If
End With
RcJournal.CloseDB
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub OpenDetailLama()
RcDetail.DBOpen " SELECT [Detail TransData].NoItem AS [Kode Barang], Inventory.ItemName AS [Nama Barang], Inventory.UOM AS Unit,  [Detail TransData].QTY_Receive AS [QTY Beli], [Detail TransData].VAT AS PPN, [Detail TransData].Price AS Harga,  ([Detail TransData].Price * ([Detail TransData].VAT / 100) + [Detail TransData].Price) * [Detail TransData].QTY_Receive AS Total" & _
                " FROM  [Detail TransData] INNER JOIN Inventory ON [Detail TransData].NoItem = Inventory.NoItem INNER JOIN TransData ON [Detail TransData].TransID = TransData.TransID WHERE     ([Detail TransData].TransID = N'" & lblFixAssets(1) & "') AND (TransData.PurchaseID = N'" & lblFixAssets(2) & "') ORDER BY [Detail TransData].NoItem", CNN
Set DGPurchase.DataSource = RcDetail.DBRecordset
End Sub

Private Function OpenPartner(ByVal Index As Integer) As Boolean
Dim RcPartner As New DBQuick
On Error GoTo xErr
mVarTmp = ""
Set mCall = New frmCaller
Select Case Index
       Case 0:
            RcPartner.DBOpen "SELECT  DateTrans AS [Tanggal Bukti], JournalID AS [No Journal], TransID AS [No Bukti], PurchaseID AS [Doc Reff] FROM [Table Journal]  WHERE     (TypeTrans = N'BKKPTP') or    (TypeTrans = N'BKMPTP') ORDER BY JournalID", CNN, lckLockReadOnly
            mCall.FromTagActive = "Dokumen Transaksi"
       Case 1:
            RcPartner.DBOpen " SELECT     [Detail TransData].NoItem AS [Kode Barang], Inventory.ItemName AS [Nama Barang], Inventory.UOM AS Unit,  [Detail TransData].QTY_Receive AS [QTY Beli], [Detail TransData].VAT AS PPN, [Detail TransData].Price AS Harga,  ([Detail TransData].Price * ([Detail TransData].VAT / 100) + [Detail TransData].Price) * [Detail TransData].QTY_Receive AS Total" & _
                             " FROM         [Detail TransData] INNER JOIN Inventory ON [Detail TransData].NoItem = Inventory.NoItem INNER JOIN TransData ON [Detail TransData].TransID = TransData.TransID WHERE     ([Detail TransData].TransID = N'" & lblFixAssets(1) & "') AND (TransData.PurchaseID = N'" & lblFixAssets(2) & "') ORDER BY [Detail TransData].NoItem", CNN, lckLockReadOnly
            mCall.FromTagActive = "Detail Transaksi"
End Select
If RcPartner.Recordcount <> 0 Then
    Set mCall.FormData = RcPartner.DBRecordset
    mVarTmp = UCase(mCall.FromTagActive)
    mCall.LookUp Me
    If mVarTmp <> "DOKUMEN TRANSAKSI" Then
       If FindOwnRecordset(MyDDE.ChildRecordset, "[Kode Barang] = '" & MyDDE.ChildRecordset.Fields(0) & "'") = True Then
          MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Kode Barang") & " Sudah Ada....! Silahkan Diulangi", "Peringatan", msgOkOnly, msgCrtical
           MyDDE.ChildRecordset.CancelBatch adAffectCurrent
           If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
          dgDetail.SetFocus
       End If
    End If
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgCrtical
   OpenPartner = True
End If
RcPartner.CloseDB
Set mCall = Nothing
Exit Function
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Function

Private Sub CancelDetailTrans()
If MyDDE.ChildRecordset.Recordcount <> 0 Then
  If Not MyDDE.ChildRecordset.EOF Then MyDDE.ChildRecordset.MoveNext
  If MyDDE.ChildRecordset.EOF And MyDDE.ChildRecordset.Recordcount > 0 Then MyDDE.ChildRecordset.MoveLast
End If
End Sub

Private Function TglIndex() As String
Dim TglHari As String
Dim TglBulan As String
Dim TglTahun As String
TglIndex = "IM-" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End Function

Private Sub TotalKas()
On Error Resume Next
Dim RcKas As New DBQuick
Dim mVarData As Variant
Dim mTotalTransData As Variant
Dim I As Long

'Init Total Dest
Set RcKas.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
mTotalTransData = 0
With RcKas.DBRecordset
     If .Recordcount <> 0 Then
        mVarData = .Getrows(.Recordcount, adBookmarkFirst, "Total")
        For I = 0 To UBound(mVarData, 2)
            mTotalTransData = mTotalTransData + IIf(Not IsNull(mVarData(0, I)), mVarData(0, I), 0)
        Next I
     End If
     LblTotalSource(1) = FormatNumber(mTotalTransData, 0)
End With
RcKas.CloseDB

'Init Total Source
'Set RcKas = New DBQuick
Set RcKas.DBRecordset = RcDetail.DBRecordset.Clone(adLockReadOnly)
mTotalTransData = 0
With RcKas.DBRecordset
     If .Recordcount <> 0 Then
        mVarData = .Getrows(.Recordcount, adBookmarkFirst, "Total")
        For I = 0 To UBound(mVarData, 2)
            mTotalTransData = mTotalTransData + IIf(Not IsNull(mVarData(0, I)), mVarData(0, I), 0)
        Next I
     End If
     LblTotalSource(0) = FormatNumber(mTotalTransData, 0)
End With
RcKas.CloseDB


End Sub

Private Sub OpenDetail(ByVal NoTransID As String, ByVal NoPartnerID As String)
Dim Rcdata As New DBQuick
If mVarAdd = False Then
   If Left(lblFixAssets(1), 2) = "RN" Then
      Rcdata.DBOpen " SELECT [Detail TransData].NoItem AS [Kode Barang], Inventory.ItemName AS [Nama Barang], Inventory.UOM AS Unit,  [Detail TransData].[QTY_IN] AS [QTY Beli], [Detail TransData].VAT AS PPN, [Detail TransData].Price AS Harga,  ([Detail TransData].Price * ([Detail TransData].VAT / 100) + [Detail TransData].Price) * [Detail TransData].QTY_Receive AS Total" & _
                    " FROM [Detail TransData] INNER JOIN Inventory ON [Detail TransData].NoItem = Inventory.NoItem INNER JOIN TransData ON [Detail TransData].TransID = TransData.TransID WHERE     (TransData.TransID = N'" & NoTransID & "') AND (TransData.PurchaseID = N'" & NoPartnerID & "') ORDER BY [Detail TransData].NoItem", CNN
   Else
      Rcdata.DBOpen " SELECT [Detail TransData].NoItem AS [Kode Barang], Inventory.ItemName AS [Nama Barang], Inventory.UOM AS Unit,  [Detail TransData].[QTY_OUT] AS [QTY Beli], [Detail TransData].VAT AS PPN, [Detail TransData].Price AS Harga,  ([Detail TransData].Price * ([Detail TransData].VAT / 100) + [Detail TransData].Price) * [Detail TransData].QTY_Receive AS Total" & _
                    " FROM [Detail TransData] INNER JOIN Inventory ON [Detail TransData].NoItem = Inventory.NoItem INNER JOIN TransData ON [Detail TransData].TransID = TransData.TransID WHERE     (TransData.TransID = N'" & NoTransID & "') AND (TransData.PurchaseID = N'" & NoPartnerID & "') ORDER BY [Detail TransData].NoItem", CNN
   End If
Else
   Rcdata.DBOpen " SELECT [Detail TransData].NoItem AS [Kode Barang], Inventory.ItemName AS [Nama Barang], Inventory.UOM AS Unit,  [Detail TransData].[QTY Adj] AS [QTY Beli], [Detail TransData].VAT AS PPN, [Detail TransData].Price AS Harga,  [Detail TransData].Price  AS Total" & _
                 " FROM [Detail TransData] INNER JOIN Inventory ON [Detail TransData].NoItem = Inventory.NoItem INNER JOIN TransData ON [Detail TransData].TransID = TransData.TransID WHERE     (TransData.TransID = N'" & NoTransID & "') AND (TransData.PurchaseID = N'" & NoPartnerID & "') ORDER BY [Detail TransData].NoItem", CNN
End If
Set MyDDE.ChildRecordset = Rcdata.DBRecordset.Clone(adLockBatchOptimistic)
Set dgDetail.DataSource = MyDDE.ChildRecordset
Rcdata.CloseDB
End Sub

Private Sub TotalTransJournal()
Dim RcTr As New DBQuick
Dim Itot As Integer
Dim mVarData As Variant
mVarKas = 0
mVarPPn = 0
mVarBarang = 0
mVarHpp = 0
Set RcTr.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
With RcTr.DBRecordset
     If .Recordcount <> 0 Then
        
        mVarData = .Getrows(.Recordcount, adBookmarkFirst)
        For Itot = 0 To UBound(mVarData, 2)
            mVarKas = mVarKas + mVarData(6, Itot)
            mVarPPn = mVarPPn + ((mVarData(4, Itot) / 100) * mVarData(5, Itot)) * mVarData(3, Itot)
            mVarHpp = mVarHpp + (HppProce(lblFixAssets(2), mVarData(0, Itot)) * mVarData(3, Itot))
        Next Itot
     End If
End With
RcTr.CloseDB
End Sub

Private Function TotalBarang() As Currency
Dim RcTr As New DBQuick
Dim Itot As Integer
Dim mVarData As Variant
Set RcTr.DBRecordset = RcDetail.DBRecordset.Clone(adLockReadOnly)
TotalBarang = 0
mVarHarga = 0
With RcTr.DBRecordset
     If .Recordcount <> 0 Then
        .Filter = "[Kode Barang] = '" & MyDDE.ChildRecordset.Fields("Kode Barang") & "'"
        If .Recordcount <> 0 Then
           TotalBarang = IIf(Not IsNull(.Fields(3)), .Fields(3), 0)
           mVarHarga = IIf(Not IsNull(.Fields(5)), .Fields(5), 0)
           
        End If
     End If
End With
RcTr.CloseDB
End Function

Private Function HppProce(ByVal NoPurchaseID As String, ByVal NoItem As String) As Double
Dim RcHpp As New DBQuick
RcHpp.DBOpen "SELECT     HPP FROM         [Detail PO] GROUP BY PurchaseID, NoItem, HPP HAVING      (PurchaseID = N'" & NoPurchaseID & "') AND (NoItem = N'" & NoItem & "')", CNN, lckLockReadOnly
With RcHpp
     If .Recordcount <> 0 Then
        HppProce = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        HppProce = 0
     End If
End With
RcHpp.CloseDB
End Function

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub
