VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{43E6F32B-2B03-46D3-9276-69426FE6D51B}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmMemo 
   Caption         =   "Memo Potongan Harga"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10050
   LinkTopic       =   "Form2"
   ScaleHeight     =   5100
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   30
      ScaleHeight     =   4305
      ScaleWidth      =   9945
      TabIndex        =   0
      Top             =   135
      Width           =   9975
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   120
         ScaleHeight     =   3705
         ScaleWidth      =   9705
         TabIndex        =   1
         Top             =   360
         Width           =   9735
         Begin VB.TextBox Text1 
            Height          =   300
            Index           =   0
            Left            =   1200
            TabIndex        =   16
            Tag             =   "memo"
            Top             =   480
            Width           =   3735
         End
         Begin VB.TextBox Text1 
            Height          =   300
            Index           =   1
            Left            =   1200
            TabIndex        =   15
            Tag             =   "memo"
            Top             =   840
            Width           =   2055
         End
         Begin VB.TextBox Text1 
            Height          =   300
            Index           =   2
            Left            =   1200
            TabIndex        =   14
            Tag             =   "memo"
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Height          =   300
            Index           =   3
            Left            =   1200
            TabIndex        =   13
            Tag             =   "memo"
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Height          =   1260
            Index           =   4
            Left            =   1200
            MultiLine       =   -1  'True
            TabIndex        =   12
            Tag             =   "memo"
            Top             =   1920
            Width           =   3735
         End
         Begin VB.TextBox Text1 
            Height          =   300
            Index           =   5
            Left            =   6600
            TabIndex        =   11
            Tag             =   "memo"
            Top             =   120
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            Height          =   300
            Index           =   6
            Left            =   6600
            TabIndex        =   10
            Tag             =   "memo"
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            Height          =   300
            Index           =   7
            Left            =   6600
            TabIndex        =   9
            Tag             =   "memo"
            Top             =   840
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            Height          =   300
            Index           =   8
            Left            =   6600
            TabIndex        =   8
            Tag             =   "memo"
            Top             =   1200
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Height          =   300
            Index           =   9
            Left            =   6600
            TabIndex        =   7
            Tag             =   "memo"
            Top             =   1560
            Width           =   2535
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   300
            Index           =   0
            Left            =   4970
            TabIndex        =   6
            Tag             =   "memo"
            Top             =   480
            Width           =   375
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   300
            Index           =   1
            Left            =   3300
            TabIndex        =   5
            Tag             =   "memo"
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   300
            Index           =   2
            Left            =   8800
            TabIndex        =   4
            Tag             =   "memo"
            Top             =   480
            Width           =   375
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   300
            Index           =   3
            Left            =   8800
            TabIndex        =   3
            Tag             =   "memo"
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   300
            Index           =   4
            Left            =   9180
            TabIndex        =   2
            Tag             =   "memo"
            Top             =   1200
            Width           =   375
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Index           =   0
            Left            =   1200
            TabIndex        =   17
            Tag             =   "memo"
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy"
            Format          =   60817411
            CurrentDate     =   39335
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Index           =   1
            Left            =   6600
            TabIndex        =   18
            Tag             =   "memo"
            Top             =   1920
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy"
            Format          =   60817411
            CurrentDate     =   39335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal"
            Height          =   255
            Index           =   0
            Left            =   -120
            TabIndex        =   30
            Top             =   150
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Customer"
            Height          =   255
            Index           =   1
            Left            =   -120
            TabIndex        =   29
            Top             =   500
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice No"
            Height          =   255
            Index           =   2
            Left            =   -120
            TabIndex        =   28
            Top             =   850
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Item No"
            Height          =   255
            Index           =   3
            Left            =   -120
            TabIndex        =   27
            Top             =   1250
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Discount"
            Height          =   255
            Index           =   4
            Left            =   -120
            TabIndex        =   26
            Top             =   1600
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Reason"
            Height          =   255
            Index           =   5
            Left            =   -120
            TabIndex        =   25
            Top             =   1950
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Lampiran"
            Height          =   255
            Index           =   6
            Left            =   5280
            TabIndex        =   24
            Top             =   140
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sales"
            Height          =   255
            Index           =   7
            Left            =   5280
            TabIndex        =   23
            Top             =   500
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Marketing"
            Height          =   255
            Index           =   8
            Left            =   5280
            TabIndex        =   22
            Top             =   850
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Document"
            Height          =   255
            Index           =   9
            Left            =   5280
            TabIndex        =   21
            Top             =   1250
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Revisi"
            Height          =   255
            Index           =   10
            Left            =   5280
            TabIndex        =   20
            Top             =   1600
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Effective Date"
            Height          =   255
            Index           =   11
            Left            =   5280
            TabIndex        =   19
            Top             =   1950
            Width           =   1095
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   31
      Top             =   4530
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   1005
      BindFormTAG     =   "memo"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmMemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
