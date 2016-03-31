VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D7BB8F75-AC9E-4E80-A526-70EA20ACFD16}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmPextract 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EXTRACTION"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12270
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   12270
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      ForeColor       =   &H80000008&
      Height          =   7125
      Left            =   0
      ScaleHeight     =   7095
      ScaleWidth      =   12255
      TabIndex        =   1
      Top             =   0
      Width           =   12285
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Air"
         ForeColor       =   &H80000008&
         Height          =   4875
         Index           =   0
         Left            =   150
         TabIndex        =   95
         Top             =   1080
         Width           =   3240
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "a"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   2115
            TabIndex        =   108
            Tag             =   "ext"
            Top             =   405
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
            TabIndex        =   107
            Tag             =   "ext"
            Top             =   675
            Width           =   930
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   2115
            TabIndex        =   106
            Tag             =   "ext"
            Top             =   945
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
            TabIndex        =   105
            Tag             =   "ext"
            Top             =   1215
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
            TabIndex        =   104
            Tag             =   "ext"
            Top             =   1485
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
            TabIndex        =   103
            Tag             =   "ext"
            Top             =   1755
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
            TabIndex        =   102
            Tag             =   "ext"
            Top             =   2025
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
            TabIndex        =   101
            Tag             =   "ext"
            Top             =   2295
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
            TabIndex        =   100
            Tag             =   "ext"
            Top             =   2565
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
            TabIndex        =   99
            Tag             =   "ext"
            Top             =   2835
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
            TabIndex        =   98
            Tag             =   "ext"
            Top             =   3105
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
            TabIndex        =   97
            Tag             =   "ext"
            Top             =   3375
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "m"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   45
            Left            =   2115
            TabIndex        =   96
            Tag             =   "ext"
            Top             =   3645
            Width           =   495
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH Akhir"
            Height          =   255
            Index           =   13
            Left            =   135
            TabIndex        =   121
            Top             =   3675
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah                                             Liter"
            Height          =   255
            Index           =   8
            Left            =   135
            TabIndex        =   120
            Top             =   435
            Width           =   2865
         End
         Begin VB.Line Line1 
            Index           =   11
            X1              =   2160
            X2              =   120
            Y1              =   3645
            Y2              =   3645
         End
         Begin VB.Line Line1 
            Index           =   10
            X1              =   2130
            X2              =   90
            Y1              =   3105
            Y2              =   3105
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   2145
            X2              =   105
            Y1              =   2295
            Y2              =   2295
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu mulai masak"
            Height          =   255
            Index           =   3
            Left            =   150
            TabIndex        =   119
            Top             =   720
            Width           =   2865
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   2160
            X2              =   120
            Y1              =   675
            Y2              =   675
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH pada awal masak"
            Height          =   255
            Index           =   4
            Left            =   135
            TabIndex        =   118
            Top             =   990
            Width           =   2865
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   2160
            X2              =   120
            Y1              =   945
            Y2              =   945
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   2145
            X2              =   105
            Y1              =   1215
            Y2              =   1215
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH pada 50 C"
            Height          =   255
            Index           =   5
            Left            =   135
            TabIndex        =   117
            Top             =   1260
            Width           =   1065
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   2145
            X2              =   105
            Y1              =   1485
            Y2              =   1485
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH pada 80 C"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   116
            Top             =   1530
            Width           =   1065
         End
         Begin VB.Line Line1 
            Index           =   43
            X1              =   2145
            X2              =   105
            Y1              =   1755
            Y2              =   1755
         End
         Begin VB.Line Line1 
            Index           =   44
            X1              =   2145
            X2              =   105
            Y1              =   2025
            Y2              =   2025
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH pada 100 C"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   115
            Top             =   1785
            Width           =   1395
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu Masakan mendidih"
            Height          =   255
            Index           =   53
            Left            =   120
            TabIndex        =   114
            Top             =   2070
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH pada 30 min"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   113
            Top             =   2340
            Width           =   1395
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH pada 60 min"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   112
            Top             =   2610
            Width           =   1395
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH pada 90 min"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   111
            Top             =   2865
            Width           =   1395
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH pada 120 min"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   110
            Top             =   3150
            Width           =   1395
         End
         Begin VB.Line Line1 
            Index           =   9
            X1              =   2160
            X2              =   120
            Y1              =   2565
            Y2              =   2565
         End
         Begin VB.Line Line1 
            Index           =   45
            X1              =   2160
            X2              =   120
            Y1              =   2835
            Y2              =   2835
         End
         Begin VB.Line Line1 
            Index           =   46
            X1              =   2145
            X2              =   105
            Y1              =   3375
            Y2              =   3375
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu Selesai masak"
            Height          =   255
            Index           =   54
            Left            =   120
            TabIndex        =   109
            Top             =   3405
            Width           =   2820
         End
         Begin VB.Line Line1 
            Index           =   47
            X1              =   2160
            X2              =   120
            Y1              =   3915
            Y2              =   3915
         End
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "Grup"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   6660
         TabIndex        =   94
         Tag             =   "ext"
         Top             =   120
         Width           =   1695
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
         TabIndex        =   93
         Tag             =   "ext"
         Top             =   405
         Width           =   1815
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "Tangki"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   6660
         TabIndex        =   92
         Tag             =   "ext"
         Top             =   420
         Width           =   1695
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   270
         Left            =   3930
         MaskColor       =   &H000000C0&
         Picture         =   "frmPextract.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   91
         Tag             =   "SPPH"
         Top             =   420
         UseMaskColor    =   -1  'True
         Width           =   300
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Penambahan Pada waktu masak"
         ForeColor       =   &H80000008&
         Height          =   4875
         Index           =   1
         Left            =   3420
         TabIndex        =   50
         Top             =   1080
         Width           =   3105
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "z"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   13
            Left            =   1260
            TabIndex        =   70
            Tag             =   "ext"
            Top             =   2295
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "y"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   14
            Left            =   2400
            TabIndex        =   69
            Tag             =   "ext"
            Top             =   2025
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "x"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   15
            Left            =   1260
            TabIndex        =   68
            Tag             =   "ext"
            Top             =   2025
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "w"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   16
            Left            =   2400
            TabIndex        =   67
            Tag             =   "ext"
            Top             =   1755
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "v"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   17
            Left            =   1260
            TabIndex        =   66
            Tag             =   "ext"
            Top             =   1755
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "u"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   18
            Left            =   2400
            TabIndex        =   65
            Tag             =   "ext"
            Top             =   1485
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "t"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   19
            Left            =   1260
            TabIndex        =   64
            Tag             =   "ext"
            Top             =   1485
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "s"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   20
            Left            =   2400
            TabIndex        =   63
            Tag             =   "ext"
            Top             =   1215
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "r"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   21
            Left            =   1260
            TabIndex        =   62
            Tag             =   "ext"
            Top             =   1215
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "q"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   22
            Left            =   2400
            TabIndex        =   61
            Tag             =   "ext"
            Top             =   945
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "p"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   23
            Left            =   1260
            TabIndex        =   60
            Tag             =   "ext"
            Top             =   945
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "o"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   24
            Left            =   1260
            TabIndex        =   59
            Tag             =   "ext"
            Top             =   675
            Width           =   930
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "n"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   25
            Left            =   1260
            TabIndex        =   58
            Tag             =   "ext"
            Top             =   405
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "aa"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   26
            Left            =   2400
            TabIndex        =   57
            Tag             =   "ext"
            Top             =   2295
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ab"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   27
            Left            =   1260
            TabIndex        =   56
            Tag             =   "ext"
            Top             =   2565
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ac"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   29
            Left            =   1260
            TabIndex        =   55
            Tag             =   "ext"
            Top             =   2835
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ad"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   28
            Left            =   1260
            TabIndex        =   54
            Tag             =   "ext"
            Top             =   3105
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ae"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   30
            Left            =   1260
            TabIndex        =   53
            Tag             =   "ext"
            Top             =   3375
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "af"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   31
            Left            =   1260
            TabIndex        =   52
            Tag             =   "ext"
            Top             =   3645
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ag"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   32
            Left            =   2295
            TabIndex        =   51
            Tag             =   "ext"
            Top             =   3645
            Width           =   495
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu"
            Height          =   255
            Index           =   9
            Left            =   1770
            TabIndex        =   90
            Top             =   2070
            Width           =   885
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "SA-14"
            Height          =   255
            Index           =   16
            Left            =   135
            TabIndex        =   89
            Top             =   2070
            Width           =   1395
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu"
            Height          =   255
            Index           =   17
            Left            =   1770
            TabIndex        =   88
            Top             =   1800
            Width           =   1395
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "OX-01"
            Height          =   255
            Index           =   18
            Left            =   150
            TabIndex        =   87
            Top             =   1800
            Width           =   1395
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu"
            Height          =   255
            Index           =   19
            Left            =   1770
            TabIndex        =   86
            Top             =   1530
            Width           =   1395
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "SA-10"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   85
            Top             =   1530
            Width           =   975
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "SA-05"
            Height          =   255
            Index           =   22
            Left            =   120
            TabIndex        =   84
            Top             =   1260
            Width           =   1065
         End
         Begin VB.Line Line1 
            Index           =   18
            X1              =   2490
            X2              =   105
            Y1              =   1485
            Y2              =   1485
         End
         Begin VB.Line Line1 
            Index           =   19
            X1              =   2550
            X2              =   105
            Y1              =   1215
            Y2              =   1215
         End
         Begin VB.Line Line1 
            Index           =   20
            X1              =   2160
            X2              =   120
            Y1              =   945
            Y2              =   945
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "SA-09"
            Height          =   255
            Index           =   24
            Left            =   135
            TabIndex        =   83
            Top             =   990
            Width           =   2865
         End
         Begin VB.Line Line1 
            Index           =   21
            X1              =   2160
            X2              =   120
            Y1              =   675
            Y2              =   675
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu "
            Height          =   255
            Index           =   25
            Left            =   150
            TabIndex        =   82
            Top             =   720
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Air                                 Liter"
            Height          =   255
            Index           =   26
            Left            =   150
            TabIndex        =   81
            Top             =   435
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "SA-15 (flake)"
            Height          =   255
            Index           =   27
            Left            =   135
            TabIndex        =   80
            Top             =   2355
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu "
            Height          =   255
            Index           =   23
            Left            =   1785
            TabIndex        =   79
            Top             =   975
            Width           =   615
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu"
            Height          =   255
            Index           =   21
            Left            =   1770
            TabIndex        =   78
            Top             =   1245
            Width           =   1065
         End
         Begin VB.Line Line1 
            Index           =   13
            X1              =   2685
            X2              =   120
            Y1              =   1755
            Y2              =   1755
         End
         Begin VB.Line Line1 
            Index           =   12
            X1              =   2685
            X2              =   120
            Y1              =   2025
            Y2              =   2025
         End
         Begin VB.Line Line1 
            Index           =   14
            X1              =   2685
            X2              =   120
            Y1              =   2295
            Y2              =   2295
         End
         Begin VB.Line Line1 
            Index           =   15
            X1              =   2685
            X2              =   120
            Y1              =   2565
            Y2              =   2565
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu"
            Height          =   255
            Index           =   28
            Left            =   1755
            TabIndex        =   77
            Top             =   2325
            Width           =   885
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "SA-15 (Buntan)"
            Height          =   255
            Index           =   29
            Left            =   120
            TabIndex        =   76
            Top             =   2625
            Width           =   1200
         End
         Begin VB.Line Line1 
            Index           =   16
            X1              =   1710
            X2              =   120
            Y1              =   2835
            Y2              =   2835
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "SA-16                            kg"
            Height          =   255
            Index           =   31
            Left            =   120
            TabIndex        =   75
            Top             =   2895
            Width           =   2160
         End
         Begin VB.Line Line1 
            Index           =   17
            X1              =   1680
            X2              =   120
            Y1              =   3105
            Y2              =   3105
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "AC-04                            liter"
            Height          =   255
            Index           =   32
            Left            =   135
            TabIndex        =   74
            Top             =   3165
            Width           =   2280
         End
         Begin VB.Line Line1 
            Index           =   22
            X1              =   1665
            X2              =   105
            Y1              =   3375
            Y2              =   3375
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu"
            Height          =   255
            Index           =   30
            Left            =   150
            TabIndex        =   73
            Top             =   3435
            Width           =   885
         End
         Begin VB.Line Line1 
            Index           =   23
            X1              =   1650
            X2              =   90
            Y1              =   3645
            Y2              =   3645
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "FA-07"
            Height          =   255
            Index           =   33
            Left            =   120
            TabIndex        =   72
            Top             =   3690
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu"
            Height          =   255
            Index           =   34
            Left            =   1785
            TabIndex        =   71
            Top             =   3705
            Width           =   885
         End
         Begin VB.Line Line1 
            Index           =   24
            X1              =   2385
            X2              =   105
            Y1              =   3915
            Y2              =   3915
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Turun ke Tangki AGAR kotor"
         ForeColor       =   &H80000008&
         Height          =   4860
         Index           =   2
         Left            =   6570
         TabIndex        =   4
         Top             =   1095
         Width           =   5550
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "au"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   33
            Left            =   4155
            TabIndex        =   28
            Tag             =   "ext"
            Top             =   2565
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "at"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   34
            Left            =   2115
            TabIndex        =   27
            Tag             =   "ext"
            Top             =   2565
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "aab"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   35
            Left            =   4155
            TabIndex        =   26
            Tag             =   "ext"
            Top             =   2295
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ar"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   36
            Left            =   2115
            TabIndex        =   25
            Tag             =   "ext"
            Top             =   2295
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ap"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   37
            Left            =   2115
            TabIndex        =   24
            Tag             =   "ext"
            Top             =   2025
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ao"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   38
            Left            =   4155
            TabIndex        =   23
            Tag             =   "ext"
            Top             =   1755
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "an"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   39
            Left            =   2115
            TabIndex        =   22
            Tag             =   "ext"
            Top             =   1755
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "am"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   40
            Left            =   4155
            TabIndex        =   21
            Tag             =   "ext"
            Top             =   1485
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "al"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   41
            Left            =   2115
            TabIndex        =   20
            Tag             =   "ext"
            Top             =   1485
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ak"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   42
            Left            =   2115
            TabIndex        =   19
            Tag             =   "ext"
            Top             =   1215
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "aj"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   46
            Left            =   2115
            TabIndex        =   18
            Tag             =   "ext"
            Top             =   945
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ai"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   47
            Left            =   2115
            TabIndex        =   17
            Tag             =   "ext"
            Top             =   675
            Width           =   930
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ah"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   48
            Left            =   2115
            TabIndex        =   16
            Tag             =   "ext"
            Top             =   405
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "aq"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   49
            Left            =   4155
            TabIndex        =   15
            Tag             =   "ext"
            Top             =   2025
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "aw"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   50
            Left            =   4155
            TabIndex        =   14
            Tag             =   "ext"
            Top             =   2835
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "av"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   51
            Left            =   2115
            TabIndex        =   13
            Tag             =   "ext"
            Top             =   2835
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ax"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   52
            Left            =   1215
            TabIndex        =   12
            Tag             =   "ext"
            Top             =   3525
            Width           =   1320
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ay"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   53
            Left            =   3855
            TabIndex        =   11
            Tag             =   "ext"
            Top             =   3525
            Width           =   1320
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "az"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   54
            Left            =   1215
            TabIndex        =   10
            Tag             =   "ext"
            Top             =   3795
            Width           =   1320
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ba"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   55
            Left            =   3855
            TabIndex        =   9
            Tag             =   "ext"
            Top             =   3795
            Width           =   1320
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "bb"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   56
            Left            =   1215
            TabIndex        =   8
            Tag             =   "ext"
            Top             =   4065
            Width           =   1320
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "bc"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   57
            Left            =   3855
            TabIndex        =   7
            Tag             =   "ext"
            Top             =   4065
            Width           =   1320
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "bd"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   58
            Left            =   1215
            TabIndex        =   6
            Tag             =   "ext"
            Top             =   4335
            Width           =   1320
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "be"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   59
            Left            =   3855
            TabIndex        =   5
            Tag             =   "ext"
            Top             =   4335
            Width           =   1320
         End
         Begin VB.Line Line1 
            Index           =   25
            X1              =   4350
            X2              =   105
            Y1              =   3105
            Y2              =   3105
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu pada 150 menit"
            Height          =   255
            Index           =   35
            Left            =   150
            TabIndex        =   49
            Top             =   2595
            Width           =   2820
         End
         Begin VB.Line Line1 
            Index           =   26
            X1              =   2145
            X2              =   105
            Y1              =   2835
            Y2              =   2835
         End
         Begin VB.Line Line1 
            Index           =   27
            X1              =   2160
            X2              =   120
            Y1              =   2565
            Y2              =   2565
         End
         Begin VB.Line Line1 
            Index           =   28
            X1              =   2160
            X2              =   120
            Y1              =   2295
            Y2              =   2295
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH pada 120 min"
            Height          =   255
            Index           =   36
            Left            =   2685
            TabIndex        =   48
            Top             =   2310
            Width           =   1395
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu pada 120 menit"
            Height          =   255
            Index           =   37
            Left            =   135
            TabIndex        =   47
            Top             =   2340
            Width           =   1650
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu pada 90 menit"
            Height          =   255
            Index           =   38
            Left            =   120
            TabIndex        =   46
            Top             =   2055
            Width           =   1515
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH pada 60 min"
            Height          =   255
            Index           =   39
            Left            =   2715
            TabIndex        =   45
            Top             =   1785
            Width           =   1395
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu pada 60 menit"
            Height          =   255
            Index           =   40
            Left            =   120
            TabIndex        =   44
            Top             =   1800
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH pada 30 min"
            Height          =   255
            Index           =   41
            Left            =   2715
            TabIndex        =   43
            Top             =   1500
            Width           =   1395
         End
         Begin VB.Line Line1 
            Index           =   29
            X1              =   4620
            X2              =   2580
            Y1              =   1755
            Y2              =   1755
         End
         Begin VB.Line Line1 
            Index           =   30
            X1              =   2145
            X2              =   105
            Y1              =   1755
            Y2              =   1755
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu pada 30 menit"
            Height          =   255
            Index           =   42
            Left            =   120
            TabIndex        =   42
            Top             =   1545
            Width           =   1515
         End
         Begin VB.Line Line1 
            Index           =   31
            X1              =   2145
            X2              =   105
            Y1              =   1485
            Y2              =   1485
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH "
            Height          =   255
            Index           =   43
            Left            =   135
            TabIndex        =   41
            Top             =   1260
            Width           =   1065
         End
         Begin VB.Line Line1 
            Index           =   32
            X1              =   2145
            X2              =   105
            Y1              =   1215
            Y2              =   1215
         End
         Begin VB.Line Line1 
            Index           =   33
            X1              =   2160
            X2              =   120
            Y1              =   945
            Y2              =   945
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu"
            Height          =   255
            Index           =   44
            Left            =   135
            TabIndex        =   40
            Top             =   975
            Width           =   2865
         End
         Begin VB.Line Line1 
            Index           =   34
            X1              =   2160
            X2              =   120
            Y1              =   675
            Y2              =   675
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu"
            Height          =   255
            Index           =   45
            Left            =   150
            TabIndex        =   39
            Top             =   720
            Width           =   2865
         End
         Begin VB.Line Line1 
            Index           =   35
            X1              =   2145
            X2              =   105
            Y1              =   2025
            Y2              =   2025
         End
         Begin VB.Line Line1 
            Index           =   36
            X1              =   4560
            X2              =   2520
            Y1              =   2565
            Y2              =   2565
         End
         Begin VB.Line Line1 
            Index           =   37
            X1              =   4590
            X2              =   2550
            Y1              =   2835
            Y2              =   2835
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah                                             Liter"
            Height          =   255
            Index           =   46
            Left            =   135
            TabIndex        =   38
            Top             =   435
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH pada 150 min"
            Height          =   255
            Index           =   47
            Left            =   2670
            TabIndex        =   37
            Top             =   2610
            Width           =   1995
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH pada 90 min"
            Height          =   255
            Index           =   48
            Left            =   2715
            TabIndex        =   36
            Top             =   2055
            Width           =   1395
         End
         Begin VB.Line Line1 
            Index           =   38
            X1              =   4590
            X2              =   2550
            Y1              =   2025
            Y2              =   2025
         End
         Begin VB.Line Line1 
            Index           =   39
            X1              =   4560
            X2              =   2520
            Y1              =   2295
            Y2              =   2295
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "suhu pada 180 menit"
            Height          =   255
            Index           =   49
            Left            =   150
            TabIndex        =   35
            Top             =   2880
            Width           =   2820
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "pH pada 180 min"
            Height          =   255
            Index           =   50
            Left            =   2670
            TabIndex        =   34
            Top             =   2865
            Width           =   1995
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Steam"
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
            Index           =   51
            Left            =   105
            TabIndex        =   33
            Top             =   3225
            Width           =   705
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu Buka                         Waktu Tutup"
            Height          =   255
            Index           =   52
            Left            =   105
            TabIndex        =   32
            Top             =   3570
            Width           =   4560
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu Buka                         Waktu Tutup"
            Height          =   255
            Index           =   55
            Left            =   120
            TabIndex        =   31
            Top             =   3840
            Width           =   4560
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu Buka                         Waktu Tutup"
            Height          =   255
            Index           =   56
            Left            =   120
            TabIndex        =   30
            Top             =   4110
            Width           =   4560
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu Buka                         Waktu Tutup"
            Height          =   255
            Index           =   57
            Left            =   120
            TabIndex        =   29
            Top             =   4380
            Width           =   4560
         End
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "desk_ext"
         Enabled         =   0   'False
         Height          =   780
         Index           =   60
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Tag             =   "ext"
         Top             =   6195
         Width           =   3885
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "id_ext"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   315
         Index           =   61
         Left            =   2100
         TabIndex        =   2
         Tag             =   "ext"
         Top             =   105
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "tgl_ext"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   0
         Left            =   2085
         TabIndex        =   122
         Tag             =   "ext"
         Top             =   690
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   58916867
         CurrentDate     =   39365
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "No Ekstrasi"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   128
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   127
         Top             =   765
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         Height          =   255
         Index           =   2
         Left            =   4770
         TabIndex        =   126
         Top             =   180
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tangki"
         Height          =   255
         Index           =   12
         Left            =   4755
         TabIndex        =   125
         Top             =   480
         Width           =   2055
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2205
         X2              =   165
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2190
         X2              =   150
         Y1              =   990
         Y2              =   990
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   6795
         X2              =   4755
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   6810
         X2              =   4770
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan :"
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
         Left            =   135
         TabIndex        =   124
         Top             =   5970
         Width           =   3150
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Extraction"
         Height          =   255
         Index           =   59
         Left            =   165
         TabIndex        =   123
         Top             =   165
         Width           =   2055
      End
   End
   Begin SemeruDC.SemeruOleDC dde 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Tag             =   "ext"
      Top             =   7095
      Width           =   12270
      _ExtentX        =   21643
      _ExtentY        =   1005
      BindFormTAG     =   "ext"
   End
End
Attribute VB_Name = "frmPextract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents Bleacing  As frmCaller
Attribute Bleacing.VB_VarHelpID = -1
Dim rsbleacing As New DBQuick

Private Sub Bleacing_RowColChange(ByVal TagForm As String, _
                                  ByVal pRecordset As ADODB.Recordset)
    txt(0).Text = rsbleacing.DBRecordset.Fields("no_ekstrasi")
End Sub

Private Sub cmdLink_Click()
    rsbleacing.DBOpen "select * from BLEACHING", CNN
    'If rsbleacing.DBRecordset.EOF Then
    'rsbleacing.DBOpen "select * from ACID_TREATMEN", CNN
    'Else
    'rsbleacing.DBOpen "select * from BLEACHING, ACID_TREATMEN where BLEACHING.no_ekstrasi <> ACID_TREATMEN.no_ekstrasi ", CNN
    'End If
    Set Bleacing = New frmCaller
    Set Bleacing.FormData = rsbleacing.DBRecordset
    Bleacing.FromTagActive = "BLEACHING TREATMEN"
    Bleacing.CaptionLink = "BLEACHING TREATMEN"
End Sub

Private Sub DDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb

        Case tmbAddNew
            CmdLink.Enabled = True
            txt(0).Enabled = True
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
        .BindFormTAG = "ext"
        Set .ActiveConnection = CNN
        .PrepareQuery = "select * from EXTRACTION"
    End With

    seting Me
    HiasFormManTell Picture2, Me
End Sub

Private Function IndexAuto() As String
    Dim Rc As New DBQuick
    Dim TglSaiki As String
    Dim Inom As Long
    TglSaiki = Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2)
    Rc.DBOpen "SELECT MAX(RIGHT(id_ext, 5)) AS MaxNom FROM [EXTRACTION] WHERE (GETDATE() = { fn NOW() })", CNN, lckLockReadOnly

    With Rc

        If .DBRecordset.Recordcount <> 0 Then
            Inom = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + 1
        Else
            Inom = 1
        End If

        Select Case Len(Trim(Str(Inom)))

            Case 0: IndexAuto = "EX/" & TglSaiki & "-" & Trim(Str(Inom))

            Case 1: IndexAuto = "EX/" & TglSaiki & "-" & "0000" & Trim(Str(Inom))

            Case 2: IndexAuto = "EX/" & TglSaiki & "-" & "000" & Trim(Str(Inom))

            Case 3: IndexAuto = "EX/" & TglSaiki & "-" & "00" & Trim(Str(Inom))

            Case 4: IndexAuto = "EX/" & TglSaiki & "-" & "0" & Trim(Str(Inom))
        End Select

    End With

End Function

Private Sub txt_GotFocus(Index As Integer)
    txt(Index).BackColor = &H79BCFF
End Sub

Private Sub txt_LostFocus(Index As Integer)
    txt(Index).BackColor = vbWhite
End Sub

Function delete()
    DDE.PrepareDelete = "delete from EXTRACTION where no_ekstrasi = '" & txt(61).Text & "'"
End Function

Function simpan()
 
    DDE.PrepareAppend = "insert into EXTRACTION (id_ext, no_ekstrasi,tgl_ext,grup,tangki,a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z,aa,ab,ac,ad,ae,af,ag,ah,ai,aj,ak,al,am,an,ao,ap,aq,ar,aab,at,au,av,aw,ax,ay,az,ba,bb,bc,bd,be,desk_ext) values " & _
       " ('" & txt(61).Text & "', '" & txt(0).Text & "', '" & Format(DTPicker1(0).value, "yyyy-MM-dd") & "', '" & DDE.GetFieldByName("grup") & "', '" & DDE.GetFieldByName("tangki") & "', '" & DDE.GetFieldByName("a") & "', '" & DDE.GetFieldByName("b") & "', '" & DDE.GetFieldByName("c") & "', '" & DDE.GetFieldByName("d") & "', '" & DDE.GetFieldByName("e") & "', '" & DDE.GetFieldByName("f") & "', '" & DDE.GetFieldByName("g") & "', " & _
       " '" & DDE.GetFieldByName("h") & "', '" & DDE.GetFieldByName("i") & "','" & DDE.GetFieldByName("j") & "', '" & DDE.GetFieldByName("k") & "', '" & DDE.GetFieldByName("l") & "', '" & DDE.GetFieldByName("m") & "', '" & DDE.GetFieldByName("n") & "', '" & DDE.GetFieldByName("o") & "', '" & DDE.GetFieldByName("p") & "', '" & DDE.GetFieldByName("q") & "', '" & DDE.GetFieldByName("r") & "', '" & DDE.GetFieldByName("s") & "', '" & DDE.GetFieldByName("t") & "', '" & DDE.GetFieldByName("u") & "', '" & DDE.GetFieldByName("v") & "', '" & DDE.GetFieldByName("w") & "' " & _
       " , '" & DDE.GetFieldByName("x") & "', '" & DDE.GetFieldByName("y") & "', '" & DDE.GetFieldByName("z") & "', '" & DDE.GetFieldByName("aa") & "', '" & DDE.GetFieldByName("ab") & "', '" & DDE.GetFieldByName("ac") & "', '" & DDE.GetFieldByName("ad") & "', '" & DDE.GetFieldByName("ae") & "', '" & DDE.GetFieldByName("af") & "', '" & DDE.GetFieldByName("ag") & "', '" & DDE.GetFieldByName("ah") & "', '" & DDE.GetFieldByName("ai") & "', '" & DDE.GetFieldByName("aj") & "', '" & DDE.GetFieldByName("ak") & "', '" & DDE.GetFieldByName("al") & "', '" & DDE.GetFieldByName("am") & "', '" & DDE.GetFieldByName("an") & "', '" & DDE.GetFieldByName("ao") & "', '" & DDE.GetFieldByName("ap") & "', '" & DDE.GetFieldByName("aq") & "', '" & DDE.GetFieldByName("ar") & "', '" & DDE.GetFieldByName("aab") & "', '" & DDE.GetFieldByName("at") & "', '" & DDE.GetFieldByName("au") & "', '" & DDE.GetFieldByName("av") & "', '" & DDE.GetFieldByName("aw") & "', '" & DDE.GetFieldByName("ax") & " ' " & _
       " , '" & DDE.GetFieldByName("ay") & "', '" & DDE.GetFieldByName("az") & "', '" & DDE.GetFieldByName("ba") & "', '" & DDE.GetFieldByName("bb") & "', '" & DDE.GetFieldByName("bc") & "', '" & DDE.GetFieldByName("bd") & "', '" & DDE.GetFieldByName("be") & "', '" & DDE.GetFieldByName("desk_ext") & "')"
                   
    DDE.PrepareUpdate = "update EXTRACTION set tgl_ext = '" & Format(DTPicker1(0).value, "yyyy-MM-dd") & "' ,Grup = '" & DDE.GetFieldByName("Grup") & "', Tangki = '" & DDE.GetFieldByName("Tangki") & "', a = '" & DDE.GetFieldByName("a") & "', b = '" & DDE.GetFieldByName("b") & "', " & _
       " c = '" & DDE.GetFieldByName("c") & "', d = '" & DDE.GetFieldByName("d") & "', e = '" & DDE.GetFieldByName("e") & "', f = '" & DDE.GetFieldByName("f") & "', g = '" & DDE.GetFieldByName("g") & "', h = '" & DDE.GetFieldByName("h") & "', i = '" & DDE.GetFieldByName("i") & "', j = '" & DDE.GetFieldByName("j") & "', k = '" & DDE.GetFieldByName("k") & "', l = '" & DDE.GetFieldByName("l") & "', " & _
       " m = '" & DDE.GetFieldByName("m") & "', n = '" & DDE.GetFieldByName("n") & "', o = '" & DDE.GetFieldByName("o") & "', p = '" & DDE.GetFieldByName("p") & "', q = '" & DDE.GetFieldByName("q") & "', r = '" & DDE.GetFieldByName("r") & "', " & _
       " s = '" & DDE.GetFieldByName("s") & "', t = '" & DDE.GetFieldByName("t") & "', u = '" & DDE.GetFieldByName("u") & "',v = '" & DDE.GetFieldByName("v") & "', w = '" & DDE.GetFieldByName("w") & "', x = '" & DDE.GetFieldByName("x") & "', y = '" & DDE.GetFieldByName("y") & "', z = '" & DDE.GetFieldByName("z") & "',aa = '" & DDE.GetFieldByName("aa") & "', ab = '" & DDE.GetFieldByName("ab") & "', ac = '" & DDE.GetFieldByName("ac") & "', ad = '" & DDE.GetFieldByName("ad") & "', ae = '" & DDE.GetFieldByName("ae") & "', af = '" & DDE.GetFieldByName("af") & "', ag = '" & DDE.GetFieldByName("ag") & "',ah = '" & DDE.GetFieldByName("ah") & "', ai = '" & DDE.GetFieldByName("ai") & "', aj = '" & DDE.GetFieldByName("aj") & "', ak = '" & DDE.GetFieldByName("ak") & "', al = '" & DDE.GetFieldByName("al") & "', am = '" & DDE.GetFieldByName("am") & "', an = '" & DDE.GetFieldByName("an") & "', ao = '" & DDE.GetFieldByName("ao") & "', ap = '" & DDE.GetFieldByName("ap") & "', " & _
       " aq = '" & DDE.GetFieldByName("aq") & "', ar = '" & DDE.GetFieldByName("ar") & "', aab = '" & DDE.GetFieldByName("aab") & "', at = '" & DDE.GetFieldByName("at") & "', au = '" & DDE.GetFieldByName("au") & "', av = '" & DDE.GetFieldByName("av") & "', aw = '" & DDE.GetFieldByName("aw") & "',ax = '" & DDE.GetFieldByName("ax") & "', ay = '" & DDE.GetFieldByName("ay") & "', az = '" & DDE.GetFieldByName("az") & "', ba = '" & DDE.GetFieldByName("ba") & "', bb = '" & DDE.GetFieldByName("bb") & "', " & _
       " bc = '" & DDE.GetFieldByName("bc") & "', bd = '" & DDE.GetFieldByName("bd") & "', be = '" & DDE.GetFieldByName("be") & "', " & _
       " desk_ext = '" & DDE.GetFieldByName("desk_ext") & "' where no_ekstrasi = '" & txt(0).Text & "'"
                   
End Function
