VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{73779082-7BF1-482D-A01F-0D9823B548F1}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmPGell 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GELLIFICATION"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12255
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   12255
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      ForeColor       =   &H80000008&
      Height          =   7500
      Left            =   0
      ScaleHeight     =   7470
      ScaleWidth      =   12255
      TabIndex        =   1
      Top             =   0
      Width           =   12285
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "id_gell"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   315
         Index           =   61
         Left            =   2100
         TabIndex        =   116
         Tag             =   "gell"
         Top             =   105
         Width           =   1815
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "desk_gell"
         Enabled         =   0   'False
         Height          =   645
         Index           =   60
         Left            =   135
         MultiLine       =   -1  'True
         TabIndex        =   115
         Tag             =   "gell"
         Top             =   6675
         Width           =   3240
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Cooling System"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3075
         Index           =   2
         Left            =   3480
         TabIndex        =   92
         Top             =   3345
         Width           =   3105
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "v"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   48
            Left            =   2115
            TabIndex        =   102
            Tag             =   "gell"
            Top             =   210
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "w"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   47
            Left            =   2115
            TabIndex        =   101
            Tag             =   "gell"
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "x"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   46
            Left            =   2115
            TabIndex        =   100
            Tag             =   "gell"
            Top             =   750
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "y"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   17
            Left            =   2115
            TabIndex        =   99
            Tag             =   "gell"
            Top             =   1020
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "z"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   18
            Left            =   2115
            TabIndex        =   98
            Tag             =   "gell"
            Top             =   1290
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "aa"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   19
            Left            =   2115
            TabIndex        =   97
            Tag             =   "gell"
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ab"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   20
            Left            =   2115
            TabIndex        =   96
            Tag             =   "gell"
            Top             =   1830
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ac"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   21
            Left            =   2115
            TabIndex        =   95
            Tag             =   "gell"
            Top             =   2100
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ad"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   22
            Left            =   2115
            TabIndex        =   94
            Tag             =   "gell"
            Top             =   2370
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ae"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   23
            Left            =   2115
            TabIndex        =   93
            Tag             =   "gell"
            Top             =   2640
            Width           =   495
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "waktu mulai"
            Height          =   255
            Index           =   46
            Left            =   120
            TabIndex        =   114
            Top             =   240
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   45
            Left            =   150
            TabIndex        =   113
            Top             =   525
            Width           =   2865
         End
         Begin VB.Line Line1 
            Index           =   34
            X1              =   2160
            X2              =   120
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line1 
            Index           =   33
            X1              =   2160
            X2              =   120
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Line Line1 
            Index           =   32
            X1              =   2160
            X2              =   120
            Y1              =   750
            Y2              =   750
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu masuk (awal)"
            Height          =   255
            Index           =   21
            Left            =   120
            TabIndex        =   112
            Top             =   510
            Width           =   1515
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu keluar (awal)"
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   111
            Top             =   780
            Width           =   1515
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH(awal)"
            Height          =   255
            Index           =   24
            Left            =   135
            TabIndex        =   110
            Top             =   1050
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   27
            Left            =   165
            TabIndex        =   109
            Top             =   1335
            Width           =   2865
         End
         Begin VB.Line Line1 
            Index           =   16
            X1              =   2175
            X2              =   135
            Y1              =   1290
            Y2              =   1290
         End
         Begin VB.Line Line1 
            Index           =   17
            X1              =   2175
            X2              =   135
            Y1              =   1830
            Y2              =   1830
         End
         Begin VB.Line Line1 
            Index           =   18
            X1              =   2175
            X2              =   135
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu masuk pada 30 min"
            Height          =   255
            Index           =   28
            Left            =   105
            TabIndex        =   108
            Top             =   1335
            Width           =   1950
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu keluar pada 30 min"
            Height          =   255
            Index           =   29
            Left            =   105
            TabIndex        =   107
            Top             =   1620
            Width           =   1950
         End
         Begin VB.Line Line1 
            Index           =   19
            X1              =   2175
            X2              =   135
            Y1              =   2370
            Y2              =   2370
         End
         Begin VB.Line Line1 
            Index           =   21
            X1              =   2175
            X2              =   135
            Y1              =   2100
            Y2              =   2100
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH pada 30 min"
            Height          =   255
            Index           =   30
            Left            =   120
            TabIndex        =   106
            Top             =   1875
            Width           =   1950
         End
         Begin VB.Line Line1 
            Index           =   22
            X1              =   2175
            X2              =   135
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Line Line1 
            Index           =   23
            X1              =   2175
            X2              =   135
            Y1              =   2910
            Y2              =   2910
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu masuk pada 60 min"
            Height          =   255
            Index           =   31
            Left            =   105
            TabIndex        =   105
            Top             =   2130
            Width           =   1950
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu keluar pada 60 min"
            Height          =   255
            Index           =   32
            Left            =   105
            TabIndex        =   104
            Top             =   2415
            Width           =   1950
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH pada 60 min"
            Height          =   255
            Index           =   33
            Left            =   120
            TabIndex        =   103
            Top             =   2670
            Width           =   1950
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   1875
         Index           =   1
         Left            =   3480
         TabIndex        =   77
         Top             =   1425
         Width           =   3105
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "q"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   25
            Left            =   1905
            TabIndex        =   82
            Tag             =   "gell"
            Top             =   375
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "r"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   13
            Left            =   1905
            TabIndex        =   81
            Tag             =   "gell"
            Top             =   645
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "s"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   14
            Left            =   1905
            TabIndex        =   80
            Tag             =   "gell"
            Top             =   915
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "t"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   15
            Left            =   1905
            TabIndex        =   79
            Tag             =   "gell"
            Top             =   1185
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "u"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   16
            Left            =   1905
            TabIndex        =   78
            Tag             =   "gell"
            Top             =   1455
            Width           =   495
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "OX-05"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   26
            Left            =   105
            TabIndex        =   91
            Top             =   390
            Width           =   615
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "jumlah                       ml"
            Height          =   255
            Index           =   25
            Left            =   975
            TabIndex        =   90
            Top             =   375
            Width           =   1725
         End
         Begin VB.Line Line1 
            Index           =   20
            X1              =   2145
            X2              =   105
            Y1              =   645
            Y2              =   645
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Air"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   9
            Left            =   105
            TabIndex        =   89
            Top             =   660
            Width           =   615
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "jumlah                       liter"
            Height          =   255
            Index           =   16
            Left            =   975
            TabIndex        =   88
            Top             =   675
            Width           =   1875
         End
         Begin VB.Line Line1 
            Index           =   12
            X1              =   2145
            X2              =   105
            Y1              =   915
            Y2              =   915
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "BA-11"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   87
            Top             =   945
            Width           =   615
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "jumlah                       liter"
            Height          =   255
            Index           =   18
            Left            =   975
            TabIndex        =   86
            Top             =   945
            Width           =   1875
         End
         Begin VB.Line Line1 
            Index           =   13
            X1              =   2160
            X2              =   120
            Y1              =   1185
            Y2              =   1185
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Air"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   85
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "jumlah                       liter"
            Height          =   255
            Index           =   20
            Left            =   975
            TabIndex        =   84
            Top             =   1215
            Width           =   1875
         End
         Begin VB.Line Line1 
            Index           =   14
            X1              =   2160
            X2              =   120
            Y1              =   1455
            Y2              =   1455
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Total pemberian Formula"
            Height          =   255
            Index           =   22
            Left            =   105
            TabIndex        =   83
            Top             =   1515
            Width           =   1875
         End
         Begin VB.Line Line1 
            Index           =   15
            X1              =   2145
            X2              =   105
            Y1              =   1725
            Y2              =   1725
         End
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   270
         Left            =   3930
         MaskColor       =   &H000000C0&
         Picture         =   "FrmPGell.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   76
         Tag             =   "SPPH"
         Top             =   420
         UseMaskColor    =   -1  'True
         Width           =   300
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "No_ekstrasi"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   0
         Left            =   2100
         TabIndex        =   75
         Tag             =   "gell"
         Top             =   405
         Width           =   1815
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "Grup"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   2085
         TabIndex        =   74
         Tag             =   "gell"
         Top             =   1005
         Width           =   1830
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Plate Heat Exchanger"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4995
         Index           =   0
         Left            =   135
         TabIndex        =   41
         Top             =   1425
         Width           =   3240
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "m"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   45
            Left            =   2115
            TabIndex        =   57
            Tag             =   "gell"
            Top             =   3645
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "l"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   44
            Left            =   2115
            TabIndex        =   56
            Tag             =   "gell"
            Top             =   3375
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "k"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   43
            Left            =   2115
            TabIndex        =   55
            Tag             =   "gell"
            Top             =   3105
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "j"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   12
            Left            =   2115
            TabIndex        =   54
            Tag             =   "gell"
            Top             =   2835
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "i"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   11
            Left            =   2115
            TabIndex        =   53
            Tag             =   "gell"
            Top             =   2565
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "h"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   10
            Left            =   2115
            TabIndex        =   52
            Tag             =   "gell"
            Top             =   2295
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "g"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   2115
            TabIndex        =   51
            Tag             =   "gell"
            Top             =   2025
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "f"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   2115
            TabIndex        =   50
            Tag             =   "gell"
            Top             =   1755
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "e"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   2115
            TabIndex        =   49
            Tag             =   "gell"
            Top             =   1485
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "d"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   2115
            TabIndex        =   48
            Tag             =   "gell"
            Top             =   1215
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   2115
            TabIndex        =   47
            Tag             =   "gell"
            Top             =   945
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "b"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   2115
            TabIndex        =   46
            Tag             =   "gell"
            Top             =   675
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "a"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   2115
            TabIndex        =   45
            Tag             =   "gell"
            Top             =   405
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "o"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2115
            TabIndex        =   44
            Tag             =   "gell"
            Top             =   4185
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "n"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   62
            Left            =   2115
            TabIndex        =   43
            Tag             =   "gell"
            Top             =   3915
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "p"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   63
            Left            =   2115
            TabIndex        =   42
            Tag             =   "gell"
            Top             =   4455
            Width           =   495
         End
         Begin VB.Line Line1 
            Index           =   47
            X1              =   2160
            X2              =   120
            Y1              =   3915
            Y2              =   3915
         End
         Begin VB.Line Line1 
            Index           =   46
            X1              =   2145
            X2              =   105
            Y1              =   3375
            Y2              =   3375
         End
         Begin VB.Line Line1 
            Index           =   45
            X1              =   2160
            X2              =   120
            Y1              =   2835
            Y2              =   2835
         End
         Begin VB.Line Line1 
            Index           =   9
            X1              =   2160
            X2              =   120
            Y1              =   2565
            Y2              =   2565
         End
         Begin VB.Line Line1 
            Index           =   44
            X1              =   2145
            X2              =   105
            Y1              =   2025
            Y2              =   2025
         End
         Begin VB.Line Line1 
            Index           =   43
            X1              =   2145
            X2              =   105
            Y1              =   1755
            Y2              =   1755
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   2145
            X2              =   105
            Y1              =   1485
            Y2              =   1485
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   2145
            X2              =   105
            Y1              =   1215
            Y2              =   1215
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   2160
            X2              =   120
            Y1              =   945
            Y2              =   945
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   2160
            X2              =   120
            Y1              =   675
            Y2              =   675
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   2145
            X2              =   105
            Y1              =   2295
            Y2              =   2295
         End
         Begin VB.Line Line1 
            Index           =   10
            X1              =   2130
            X2              =   90
            Y1              =   3105
            Y2              =   3105
         End
         Begin VB.Line Line1 
            Index           =   11
            X1              =   2160
            X2              =   120
            Y1              =   3645
            Y2              =   3645
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu Mulai"
            Height          =   255
            Index           =   3
            Left            =   165
            TabIndex        =   73
            Top             =   405
            Width           =   1380
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu masuk (awal)"
            Height          =   255
            Index           =   4
            Left            =   165
            TabIndex        =   72
            Top             =   705
            Width           =   1515
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu keluar (awal)"
            Height          =   255
            Index           =   5
            Left            =   165
            TabIndex        =   71
            Top             =   990
            Width           =   1515
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu masuk pada 30 min"
            Height          =   255
            Index           =   6
            Left            =   150
            TabIndex        =   70
            Top             =   1245
            Width           =   1950
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu keluar pada 30 min"
            Height          =   255
            Index           =   7
            Left            =   150
            TabIndex        =   69
            Top             =   1530
            Width           =   1950
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu masuk pada 60 min"
            Height          =   255
            Index           =   8
            Left            =   135
            TabIndex        =   68
            Top             =   1800
            Width           =   1950
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu keluar pada 60 min"
            Height          =   255
            Index           =   10
            Left            =   150
            TabIndex        =   67
            Top             =   2085
            Width           =   1950
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu masuk pada 90 min"
            Height          =   255
            Index           =   11
            Left            =   135
            TabIndex        =   66
            Top             =   2340
            Width           =   1950
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu keluar pada 90 min"
            Height          =   255
            Index           =   12
            Left            =   135
            TabIndex        =   65
            Top             =   2625
            Width           =   1950
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu masuk pada 120 min"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   64
            Top             =   2895
            Width           =   1935
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu keluar pada 120 min    "
            Height          =   255
            Index           =   14
            Left            =   135
            TabIndex        =   63
            Top             =   3150
            Width           =   1890
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   2145
            X2              =   105
            Y1              =   4455
            Y2              =   4455
         End
         Begin VB.Line Line1 
            Index           =   40
            X1              =   2145
            X2              =   105
            Y1              =   4185
            Y2              =   4185
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu masuk pada 150 min      "
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   62
            Top             =   3405
            Width           =   1980
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu keluar pada 150 min "
            Height          =   255
            Index           =   53
            Left            =   120
            TabIndex        =   61
            Top             =   3690
            Width           =   1920
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu masuk pada 180 min "
            Height          =   255
            Index           =   54
            Left            =   105
            TabIndex        =   60
            Top             =   3960
            Width           =   2670
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu keluar pada 180 min"
            Height          =   255
            Index           =   60
            Left            =   120
            TabIndex        =   59
            Top             =   4230
            Width           =   2880
         End
         Begin VB.Line Line1 
            Index           =   41
            X1              =   2145
            X2              =   105
            Y1              =   4725
            Y2              =   4725
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "waktu selesai     "
            Height          =   255
            Index           =   61
            Left            =   120
            TabIndex        =   58
            Top             =   4500
            Width           =   1740
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Cooling System"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4995
         Index           =   3
         Left            =   6675
         TabIndex        =   2
         Top             =   1425
         Width           =   5460
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ao"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   24
            Left            =   2115
            TabIndex        =   21
            Tag             =   "gell"
            Top             =   2640
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "an"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   26
            Left            =   2115
            TabIndex        =   20
            Tag             =   "gell"
            Top             =   2370
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "am"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   27
            Left            =   2115
            TabIndex        =   19
            Tag             =   "gell"
            Top             =   2100
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "al"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   28
            Left            =   2115
            TabIndex        =   18
            Tag             =   "gell"
            Top             =   1830
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ak"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   29
            Left            =   2115
            TabIndex        =   17
            Tag             =   "gell"
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "aj"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   30
            Left            =   2115
            TabIndex        =   16
            Tag             =   "gell"
            Top             =   1290
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ai"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   31
            Left            =   2115
            TabIndex        =   15
            Tag             =   "gell"
            Top             =   1020
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ah"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   32
            Left            =   2115
            TabIndex        =   14
            Tag             =   "gell"
            Top             =   750
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ag"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   33
            Left            =   2115
            TabIndex        =   13
            Tag             =   "gell"
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "af"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   34
            Left            =   2115
            TabIndex        =   12
            Tag             =   "gell"
            Top             =   210
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "aq"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   35
            Left            =   2115
            TabIndex        =   11
            Tag             =   "gell"
            Top             =   3180
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ap"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   36
            Left            =   2115
            TabIndex        =   10
            Tag             =   "gell"
            Top             =   2910
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ar"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   37
            Left            =   2115
            TabIndex        =   9
            Tag             =   "gell"
            Top             =   3450
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "au"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   38
            Left            =   4770
            TabIndex        =   8
            Tag             =   "gell"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "at"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   39
            Left            =   4770
            TabIndex        =   7
            Tag             =   "gell"
            Top             =   450
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "aab"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   40
            Left            =   4770
            TabIndex        =   6
            Tag             =   "gell"
            Top             =   180
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "aw"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   41
            Left            =   4770
            TabIndex        =   5
            Tag             =   "gell"
            Top             =   1260
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "av"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   42
            Left            =   4770
            TabIndex        =   4
            Tag             =   "gell"
            Top             =   990
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ax"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   49
            Left            =   4770
            TabIndex        =   3
            Tag             =   "gell"
            Top             =   1530
            Width           =   495
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH pada 150 min"
            Height          =   255
            Index           =   34
            Left            =   120
            TabIndex        =   40
            Top             =   2415
            Width           =   1950
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu keluar pada 150 min"
            Height          =   255
            Index           =   35
            Left            =   105
            TabIndex        =   39
            Top             =   2160
            Width           =   1950
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu masuk pada 150 min"
            Height          =   255
            Index           =   36
            Left            =   105
            TabIndex        =   38
            Top             =   1875
            Width           =   1950
         End
         Begin VB.Line Line1 
            Index           =   24
            X1              =   2175
            X2              =   135
            Y1              =   2910
            Y2              =   2910
         End
         Begin VB.Line Line1 
            Index           =   25
            X1              =   2175
            X2              =   135
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH pada 120 min"
            Height          =   255
            Index           =   37
            Left            =   120
            TabIndex        =   37
            Top             =   1605
            Width           =   1950
         End
         Begin VB.Line Line1 
            Index           =   26
            X1              =   2175
            X2              =   135
            Y1              =   2100
            Y2              =   2100
         End
         Begin VB.Line Line1 
            Index           =   27
            X1              =   2175
            X2              =   135
            Y1              =   2370
            Y2              =   2370
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu keluar pada 120 min"
            Height          =   255
            Index           =   38
            Left            =   105
            TabIndex        =   36
            Top             =   1335
            Width           =   1950
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu masuk pada 120 min"
            Height          =   255
            Index           =   39
            Left            =   105
            TabIndex        =   35
            Top             =   1050
            Width           =   1950
         End
         Begin VB.Line Line1 
            Index           =   28
            X1              =   2175
            X2              =   135
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Line Line1 
            Index           =   29
            X1              =   2175
            X2              =   135
            Y1              =   1830
            Y2              =   1830
         End
         Begin VB.Line Line1 
            Index           =   30
            X1              =   2175
            X2              =   135
            Y1              =   1290
            Y2              =   1290
         End
         Begin VB.Line Line1 
            Index           =   31
            X1              =   2160
            X2              =   120
            Y1              =   750
            Y2              =   750
         End
         Begin VB.Line Line1 
            Index           =   35
            X1              =   2160
            X2              =   120
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Line Line1 
            Index           =   36
            X1              =   2160
            X2              =   120
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH pada 90 min"
            Height          =   255
            Index           =   41
            Left            =   135
            TabIndex        =   34
            Top             =   795
            Width           =   1950
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu keluar pada 90 min"
            Height          =   255
            Index           =   42
            Left            =   120
            TabIndex        =   33
            Top             =   510
            Width           =   1950
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu masuk pada 90 min"
            Height          =   255
            Index           =   43
            Left            =   120
            TabIndex        =   32
            Top             =   210
            Width           =   1800
         End
         Begin VB.Line Line1 
            Index           =   37
            X1              =   2175
            X2              =   135
            Y1              =   3450
            Y2              =   3450
         End
         Begin VB.Line Line1 
            Index           =   38
            X1              =   2175
            X2              =   135
            Y1              =   3180
            Y2              =   3180
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH pada 180 min"
            Height          =   255
            Index           =   40
            Left            =   120
            TabIndex        =   31
            Top             =   3210
            Width           =   1950
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu keluar pada 180 min"
            Height          =   255
            Index           =   44
            Left            =   105
            TabIndex        =   30
            Top             =   2955
            Width           =   1950
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu masuk pada 180 min"
            Height          =   255
            Index           =   47
            Left            =   105
            TabIndex        =   29
            Top             =   2670
            Width           =   1950
         End
         Begin VB.Line Line1 
            Index           =   39
            X1              =   2160
            X2              =   120
            Y1              =   3720
            Y2              =   3720
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "waktu selesai     "
            Height          =   255
            Index           =   48
            Left            =   135
            TabIndex        =   28
            Top             =   3510
            Width           =   1740
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH kempu no 5"
            Height          =   255
            Index           =   49
            Left            =   2775
            TabIndex        =   27
            Top             =   495
            Width           =   1950
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH kempu no 3"
            Height          =   255
            Index           =   50
            Left            =   2760
            TabIndex        =   26
            Top             =   240
            Width           =   1950
         End
         Begin VB.Line Line1 
            Index           =   42
            X1              =   4830
            X2              =   2790
            Y1              =   990
            Y2              =   990
         End
         Begin VB.Line Line1 
            Index           =   48
            X1              =   4830
            X2              =   2790
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line1 
            Index           =   49
            X1              =   4830
            X2              =   2790
            Y1              =   450
            Y2              =   450
         End
         Begin VB.Line Line1 
            Index           =   50
            X1              =   4830
            X2              =   2790
            Y1              =   1530
            Y2              =   1530
         End
         Begin VB.Line Line1 
            Index           =   51
            X1              =   4830
            X2              =   2790
            Y1              =   1260
            Y2              =   1260
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH kempu no 15"
            Height          =   255
            Index           =   51
            Left            =   2775
            TabIndex        =   25
            Top             =   1290
            Width           =   1950
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH kempu no 12"
            Height          =   255
            Index           =   52
            Left            =   2760
            TabIndex        =   24
            Top             =   1035
            Width           =   1950
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH kempu no 9"
            Height          =   255
            Index           =   55
            Left            =   2760
            TabIndex        =   23
            Top             =   750
            Width           =   1950
         End
         Begin VB.Line Line1 
            Index           =   52
            X1              =   4815
            X2              =   2775
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Total kempu"
            Height          =   255
            Index           =   56
            Left            =   2790
            TabIndex        =   22
            Top             =   1590
            Width           =   1740
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "tgl_gell"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   0
         Left            =   2085
         TabIndex        =   117
         Tag             =   "gell"
         Top             =   690
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   60686339
         CurrentDate     =   39365
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "ID GELLIFICATION"
         Height          =   255
         Index           =   59
         Left            =   165
         TabIndex        =   122
         Top             =   165
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   58
         Left            =   150
         TabIndex        =   121
         Top             =   6450
         Width           =   3150
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   2190
         X2              =   150
         Y1              =   1305
         Y2              =   1305
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2190
         X2              =   150
         Y1              =   990
         Y2              =   990
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2205
         X2              =   165
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         Height          =   255
         Index           =   2
         Left            =   165
         TabIndex        =   120
         Top             =   1050
         Width           =   1200
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   119
         Top             =   765
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "No Ekstrasi"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   118
         Top             =   480
         Width           =   2055
      End
   End
   Begin SemeruDC.SemeruOleDC dde 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Tag             =   "gell"
      Top             =   7500
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   1005
      BindFormTAG     =   "gell"
   End
End
Attribute VB_Name = "FrmPGell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mAlkali As frmCaller
Attribute mAlkali.VB_VarHelpID = -1
Dim rsalkali As New DBQuick
Private Sub cmdLink_Click()
   rsalkali.DBOpen "select * from ACID_TREATMEN", CNN
   If rsalkali.DBRecordset.EOF Then
   rsalkali.DBOpen "select * from ALKALI_TREATMEN", CNN
   Else
   rsalkali.DBOpen "select * from ALKALI_TREATMEN, ACID_TREATMEN where ALKALI_TREATMEN.no_ekstrasi <> ACID_TREATMEN.no_ekstrasi ", CNN
   End If
   
   Set mAlkali = New frmCaller
   Set mAlkali.FormData = rsalkali.DBRecordset
   mAlkali.FromTagActive = "ALKALI TREATMEN"
   mAlkali.CaptionLink = "ALKALI TREATMEN"
End Sub

Private Sub DDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
Case tmbAddNew
     cmdLink.Enabled = True
     txt(0).Enabled = True
      txt(0).SetFocus
     txt(61).Enabled = False
     txt(61).Text = IndexAuto
End Select
End Sub

Private Sub DDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb

Case tmbSave
     DDE.IsChildMemberReady = True
     simpan
Case tmbDelete
     delete
End Select
End Sub

Private Sub Form_Load()
With DDE
 Set .BindForm = Me
     .BindFormTAG = "GELL"
 Set .ActiveConnection = CNN
     .PrepareQuery = "select * from GELLIFICATION"
End With
HiasFormManTell Picture2, Me
seting Me
End Sub

Private Function IndexAuto() As String
Dim Rc As New DBQuick
Dim TglSaiki As String
Dim Inom As Long
TglSaiki = Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2)
Rc.DBOpen "SELECT MAX(RIGHT(id_gell, 5)) AS MaxNom FROM [GELLIFICATION] WHERE (GETDATE() = { fn NOW() })", CNN, lckLockReadOnly
With Rc
     If .DBRecordset.Recordcount <> 0 Then
        Inom = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + 1
     Else
        Inom = 1
     End If
     Select Case Len(Trim(Str(Inom)))
            Case 0: IndexAuto = "GE/" & TglSaiki & "-" & Trim(Str(Inom))
            Case 1: IndexAuto = "GE/" & TglSaiki & "-" & "0000" & Trim(Str(Inom))
            Case 2: IndexAuto = "GE/" & TglSaiki & "-" & "000" & Trim(Str(Inom))
            Case 3: IndexAuto = "GE/" & TglSaiki & "-" & "00" & Trim(Str(Inom))
            Case 4: IndexAuto = "GE/" & TglSaiki & "-" & "0" & Trim(Str(Inom))
     End Select
End With
End Function

Private Sub mAlkali_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
DDE.GetFieldByName("no_ekstrasi") = rsalkali.DBRecordset.Fields("no_ekstrasi")
End Sub

Private Sub txt_GotFocus(Index As Integer)
txt(Index).BackColor = &H79BCFF
End Sub

Private Sub txt_LostFocus(Index As Integer)
txt(Index).BackColor = vbWhite
End Sub

Function delete()
DDE.PrepareDelete = "delete from GELLIFICATION where no_ekstrasi = '" & DDE.GetFieldByName("no_ekstrasi") & "'"
End Function

Function simpan()
 
DDE.PrepareAppend = "insert into GELLIFICATION (id_gell, no_ekstrasi,tgl_gell,grup,a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z,aa,ab,ac,ad,ae,af,ag,ah,ai,aj,ak,al,am,an,ao,ap,aq,ar,aab,at,au,av,aw,ax,desk_gell) values " & _
                    " ('" & txt(61).Text & "', '" & DDE.GetFieldByName("no_ekstrasi") & "', '" & Format(DTPicker1(0).Value, "yyyy-MM-dd") & "', '" & DDE.GetFieldByName("grup") & "', '" & DDE.GetFieldByName("a") & "', '" & DDE.GetFieldByName("b") & "', '" & DDE.GetFieldByName("c") & "', '" & DDE.GetFieldByName("d") & "', '" & DDE.GetFieldByName("e") & "', '" & DDE.GetFieldByName("f") & "', '" & DDE.GetFieldByName("g") & "', '" & DDE.GetFieldByName("h") & "', '" & DDE.GetFieldByName("i") & "', '" & DDE.GetFieldByName("j") & "', '" & DDE.GetFieldByName("k") & "', '" & DDE.GetFieldByName("l") & "', '" & DDE.GetFieldByName("m") & "', '" & DDE.GetFieldByName("n") & "', '" & DDE.GetFieldByName("o") & "', '" & DDE.GetFieldByName("p") & "', '" & DDE.GetFieldByName("q") & "', '" & DDE.GetFieldByName("r") & "', '" & DDE.GetFieldByName("s") & "', '" & DDE.GetFieldByName("t") & "', '" & DDE.GetFieldByName("u") & "', '" & DDE.GetFieldByName("v") & "', '" & DDE.GetFieldByName("w") & "' " & _
                    " , '" & DDE.GetFieldByName("x") & "', '" & DDE.GetFieldByName("y") & "', '" & DDE.GetFieldByName("z") & "', '" & DDE.GetFieldByName("aa") & "', '" & DDE.GetFieldByName("ab") & "', '" & DDE.GetFieldByName("ac") & "', '" & DDE.GetFieldByName("ad") & "', '" & DDE.GetFieldByName("ae") & "', '" & DDE.GetFieldByName("af") & "', '" & DDE.GetFieldByName("ag") & "', '" & DDE.GetFieldByName("ah") & "', '" & DDE.GetFieldByName("ai") & "', '" & DDE.GetFieldByName("aj") & "', '" & DDE.GetFieldByName("ak") & "', '" & DDE.GetFieldByName("al") & "', '" & DDE.GetFieldByName("am") & "', '" & DDE.GetFieldByName("an") & "', '" & DDE.GetFieldByName("ao") & "', '" & DDE.GetFieldByName("ap") & "', '" & DDE.GetFieldByName("aq") & "', '" & DDE.GetFieldByName("ar") & "', '" & DDE.GetFieldByName("aab") & "', '" & DDE.GetFieldByName("at") & "', '" & DDE.GetFieldByName("au") & "', '" & DDE.GetFieldByName("av") & "', '" & DDE.GetFieldByName("aw") & "', '" & DDE.GetFieldByName("ax") & " ' " & _
                    " , '" & DDE.GetFieldByName("desk_gell") & "')"
                    
                    
DDE.PrepareUpdate = "update GELLIFICATION set tgl_gell = '" & Format(DTPicker1(0).Value, "yyyy-MM-dd") & "', Grup = '" & DDE.GetFieldByName("Grup") & "', " & _
                " a = '" & DDE.GetFieldByName("a") & "', b = '" & DDE.GetFieldByName("b") & "', c = '" & DDE.GetFieldByName("c") & "', d = '" & DDE.GetFieldByName("d") & "', " & _
                " e = '" & DDE.GetFieldByName("e") & "', f = '" & DDE.GetFieldByName("f") & "', g = '" & DDE.GetFieldByName("g") & "', h = '" & DDE.GetFieldByName("h") & "', " & _
                " i = '" & DDE.GetFieldByName("i") & "', j = '" & DDE.GetFieldByName("j") & "', k = '" & DDE.GetFieldByName("k") & "', l = '" & DDE.GetFieldByName("l") & "', " & _
                " m = '" & DDE.GetFieldByName("m") & "', n = '" & DDE.GetFieldByName("n") & "', o = '" & DDE.GetFieldByName("o") & "', p = '" & DDE.GetFieldByName("p") & "', " & _
                " q = '" & DDE.GetFieldByName("q") & "', r = '" & DDE.GetFieldByName("r") & "', s = '" & DDE.GetFieldByName("s") & "', t = '" & DDE.GetFieldByName("t") & "', " & _
                " u = '" & DDE.GetFieldByName("u") & "', v = '" & DDE.GetFieldByName("v") & "', w = '" & DDE.GetFieldByName("w") & "', x = '" & DDE.GetFieldByName("x") & "', " & _
                " y = '" & DDE.GetFieldByName("y") & "', z = '" & DDE.GetFieldByName("z") & "', aa = '" & DDE.GetFieldByName("aa") & "', ab = '" & DDE.GetFieldByName("ab") & "', " & _
                " ac = '" & DDE.GetFieldByName("ac") & "', ad = '" & DDE.GetFieldByName("ad") & "', ae = '" & DDE.GetFieldByName("ae") & "', af = '" & DDE.GetFieldByName("af") & "', " & _
                " ag = '" & DDE.GetFieldByName("ag") & "', ah = '" & DDE.GetFieldByName("ah") & "', ai = '" & DDE.GetFieldByName("ai") & "', aj = '" & DDE.GetFieldByName("aj") & "', " & _
                " ak = '" & DDE.GetFieldByName("ak") & "', al = '" & DDE.GetFieldByName("al") & "', am = '" & DDE.GetFieldByName("am") & "', an = '" & DDE.GetFieldByName("an") & "', " & _
                " ao = '" & DDE.GetFieldByName("ao") & "', ap = '" & DDE.GetFieldByName("ap") & "', aq = '" & DDE.GetFieldByName("aq") & "', ar = '" & DDE.GetFieldByName("ar") & "', " & _
                " aab = '" & DDE.GetFieldByName("aab") & "', at = '" & DDE.GetFieldByName("at") & "', au = '" & DDE.GetFieldByName("au") & "', av = '" & DDE.GetFieldByName("av") & "', " & _
                " aw = '" & DDE.GetFieldByName("aw") & "', ax = '" & DDE.GetFieldByName("ax") & "', desk_gell = '" & DDE.GetFieldByName("desk_gell") & "' where no_ekstrasi ='" & DDE.GetFieldByName("no_ekstrasi") & "'"

End Function
