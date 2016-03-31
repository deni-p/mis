VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B1E614FF-F86D-4F68-A86F-2584A0570C66}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmAlkali 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alkali Treatment"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11925
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      ForeColor       =   &H80000008&
      Height          =   7395
      Left            =   15
      ScaleHeight     =   7365
      ScaleWidth      =   12750
      TabIndex        =   1
      Top             =   0
      Width           =   12780
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "Tangki"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   2
         Left            =   2130
         TabIndex        =   95
         Tag             =   "alkali"
         Top             =   975
         Width           =   1695
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "Grup"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   1
         Left            =   2130
         TabIndex        =   94
         Tag             =   "alkali"
         Top             =   675
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   4740
         Index           =   0
         Left            =   90
         TabIndex        =   69
         Top             =   1365
         Width           =   3720
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "No_stock"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   3
            Left            =   2070
            TabIndex        =   79
            Tag             =   "alkali"
            Top             =   255
            Width           =   1335
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "Berat"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   4
            Left            =   2070
            TabIndex        =   78
            Tag             =   "alkali"
            Top             =   555
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "JML_ALKALI_bekas"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   5
            Left            =   2070
            TabIndex        =   77
            Tag             =   "alkali"
            Top             =   1185
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "Konsentrasi_alkali_bekas"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   6
            Left            =   2070
            TabIndex        =   76
            Tag             =   "alkali"
            Top             =   1485
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "Jumlah_air"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   7
            Left            =   2070
            TabIndex        =   75
            Tag             =   "alkali"
            Top             =   2115
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "JML_ALKALI_BARU"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   8
            Left            =   2070
            TabIndex        =   74
            Tag             =   "alkali"
            Top             =   2745
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "Konsentrasi_alkali_baru"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   9
            Left            =   2070
            TabIndex        =   73
            Tag             =   "alkali"
            Top             =   3045
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "JML_alkali_akhir"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   10
            Left            =   2070
            TabIndex        =   72
            Tag             =   "alkali"
            Top             =   3690
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "konsentrasi_alkali_akhir"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   11
            Left            =   2070
            TabIndex        =   71
            Tag             =   "alkali"
            Top             =   3990
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "Suhu_alkali_akhir"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   12
            Left            =   2070
            TabIndex        =   70
            Tag             =   "alkali"
            Top             =   4290
            Width           =   495
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "No Stock Rumput Laut"
            Height          =   255
            Index           =   3
            Left            =   150
            TabIndex        =   93
            Top             =   315
            Width           =   2055
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Berat Rumput Laut                         Kg"
            Height          =   255
            Index           =   4
            Left            =   150
            TabIndex        =   92
            Top             =   645
            Width           =   2745
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Larutan Alkali Bekas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   150
            TabIndex        =   91
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah                                            Liter"
            Height          =   255
            Index           =   6
            Left            =   150
            TabIndex        =   90
            Top             =   1260
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Konsentrasi                                     %"
            Height          =   255
            Index           =   7
            Left            =   150
            TabIndex        =   89
            Top             =   1575
            Width           =   2820
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
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   9
            Left            =   150
            TabIndex        =   88
            Top             =   1890
            Width           =   2055
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Alkali Baru"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   10
            Left            =   150
            TabIndex        =   87
            Top             =   2520
            Width           =   2055
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Larutan Alkali Akhir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   15
            Left            =   150
            TabIndex        =   86
            Top             =   3480
            Width           =   2055
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Suhu                                                C"
            Height          =   255
            Index           =   18
            Left            =   150
            TabIndex        =   85
            Top             =   4395
            Width           =   2730
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   2175
            X2              =   135
            Y1              =   555
            Y2              =   555
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   2175
            X2              =   135
            Y1              =   855
            Y2              =   855
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   2160
            X2              =   120
            Y1              =   1485
            Y2              =   1485
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   2160
            X2              =   120
            Y1              =   1785
            Y2              =   1785
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   2160
            X2              =   120
            Y1              =   2415
            Y2              =   2415
         End
         Begin VB.Line Line1 
            Index           =   9
            X1              =   2190
            X2              =   150
            Y1              =   3045
            Y2              =   3045
         End
         Begin VB.Line Line1 
            Index           =   10
            X1              =   2175
            X2              =   135
            Y1              =   3345
            Y2              =   3345
         End
         Begin VB.Line Line1 
            Index           =   11
            X1              =   2175
            X2              =   135
            Y1              =   3990
            Y2              =   3990
         End
         Begin VB.Line Line1 
            Index           =   12
            X1              =   2175
            X2              =   135
            Y1              =   4290
            Y2              =   4290
         End
         Begin VB.Line Line1 
            Index           =   13
            X1              =   2175
            X2              =   135
            Y1              =   4590
            Y2              =   4590
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah                                            Liter"
            Height          =   255
            Index           =   8
            Left            =   135
            TabIndex        =   84
            Top             =   2205
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah                                            Kg"
            Height          =   255
            Index           =   11
            Left            =   165
            TabIndex        =   83
            Top             =   2820
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah                                            Liter"
            Height          =   255
            Index           =   13
            Left            =   165
            TabIndex        =   82
            Top             =   3780
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Konsentrasi                                     %"
            Height          =   255
            Index           =   16
            Left            =   150
            TabIndex        =   81
            Top             =   4065
            Width           =   2820
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Konsentrasi                                     %"
            Height          =   255
            Index           =   14
            Left            =   150
            TabIndex        =   80
            Top             =   3135
            Width           =   2820
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Waktu Mulai Treatment"
         ForeColor       =   &H80000008&
         Height          =   4740
         Index           =   1
         Left            =   3900
         TabIndex        =   42
         Top             =   1365
         Width           =   3030
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "suhu_1"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   22
            Left            =   2070
            TabIndex        =   55
            Tag             =   "alkali"
            Top             =   420
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "suhu_2"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   13
            Left            =   2070
            TabIndex        =   54
            Tag             =   "alkali"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "suhu_3"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   14
            Left            =   2070
            TabIndex        =   53
            Tag             =   "alkali"
            Top             =   1020
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "suhu_4"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   15
            Left            =   2070
            TabIndex        =   52
            Tag             =   "alkali"
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "suhu_5"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   16
            Left            =   2070
            TabIndex        =   51
            Tag             =   "alkali"
            Top             =   1620
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "suhu_6"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   17
            Left            =   2070
            TabIndex        =   50
            Tag             =   "alkali"
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "suhu_7"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   18
            Left            =   2070
            TabIndex        =   49
            Tag             =   "alkali"
            Top             =   2220
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "suhu_8"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   19
            Left            =   2070
            TabIndex        =   48
            Tag             =   "alkali"
            Top             =   2520
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "suhu_9"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   20
            Left            =   2070
            TabIndex        =   47
            Tag             =   "alkali"
            Top             =   2820
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "suhu_10"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   21
            Left            =   2070
            TabIndex        =   46
            Tag             =   "alkali"
            Top             =   3120
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "suhu_11"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   23
            Left            =   2070
            TabIndex        =   45
            Tag             =   "alkali"
            Top             =   3420
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "suhu_12"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   24
            Left            =   2070
            TabIndex        =   44
            Tag             =   "alkali"
            Top             =   3720
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "waktu_selesai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   41
            Left            =   2070
            TabIndex        =   43
            Tag             =   "alkali"
            Top             =   4020
            Width           =   495
         End
         Begin VB.Line Line1 
            Index           =   23
            X1              =   2175
            X2              =   135
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Suhu Setelah 1 jam               C"
            Height          =   255
            Index           =   31
            Left            =   135
            TabIndex        =   68
            Top             =   480
            Width           =   2715
         End
         Begin VB.Line Line1 
            Index           =   14
            X1              =   2175
            X2              =   135
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Suhu Setelah 2 jam               C"
            Height          =   255
            Index           =   17
            Left            =   150
            TabIndex        =   67
            Top             =   780
            Width           =   2715
         End
         Begin VB.Line Line1 
            Index           =   15
            X1              =   2175
            X2              =   135
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Suhu Setelah 3 jam               C"
            Height          =   255
            Index           =   19
            Left            =   150
            TabIndex        =   66
            Top             =   1080
            Width           =   2715
         End
         Begin VB.Line Line1 
            Index           =   16
            X1              =   2175
            X2              =   135
            Y1              =   1620
            Y2              =   1620
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Suhu Setelah 4 jam               C"
            Height          =   255
            Index           =   20
            Left            =   150
            TabIndex        =   65
            Top             =   1380
            Width           =   2715
         End
         Begin VB.Line Line1 
            Index           =   17
            X1              =   2175
            X2              =   135
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Suhu Setelah 5 jam               C"
            Height          =   255
            Index           =   21
            Left            =   150
            TabIndex        =   64
            Top             =   1680
            Width           =   2715
         End
         Begin VB.Line Line1 
            Index           =   18
            X1              =   2175
            X2              =   135
            Y1              =   2220
            Y2              =   2220
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Suhu Setelah 6 jam               C"
            Height          =   255
            Index           =   22
            Left            =   150
            TabIndex        =   63
            Top             =   1980
            Width           =   2715
         End
         Begin VB.Line Line1 
            Index           =   19
            X1              =   2175
            X2              =   135
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Suhu Setelah 7 jam               C"
            Height          =   255
            Index           =   23
            Left            =   150
            TabIndex        =   62
            Top             =   2280
            Width           =   2715
         End
         Begin VB.Line Line1 
            Index           =   20
            X1              =   2175
            X2              =   135
            Y1              =   2820
            Y2              =   2820
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Suhu Setelah 8 jam               C"
            Height          =   255
            Index           =   24
            Left            =   150
            TabIndex        =   61
            Top             =   2580
            Width           =   2715
         End
         Begin VB.Line Line1 
            Index           =   21
            X1              =   2175
            X2              =   135
            Y1              =   3120
            Y2              =   3120
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Suhu Setelah 9 jam               C"
            Height          =   255
            Index           =   25
            Left            =   150
            TabIndex        =   60
            Top             =   2880
            Width           =   2715
         End
         Begin VB.Line Line1 
            Index           =   22
            X1              =   2175
            X2              =   135
            Y1              =   3420
            Y2              =   3420
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Suhu Setelah 10  jam            C"
            Height          =   255
            Index           =   26
            Left            =   150
            TabIndex        =   59
            Top             =   3180
            Width           =   2850
         End
         Begin VB.Line Line1 
            Index           =   24
            X1              =   2175
            X2              =   135
            Y1              =   3720
            Y2              =   3720
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Suhu Setelah 11  jam            C"
            Height          =   255
            Index           =   27
            Left            =   150
            TabIndex        =   58
            Top             =   3480
            Width           =   2850
         End
         Begin VB.Line Line1 
            Index           =   25
            X1              =   2175
            X2              =   135
            Y1              =   4020
            Y2              =   4020
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Suhu Setelah 12  jam            C          "
            Height          =   255
            Index           =   28
            Left            =   150
            TabIndex        =   57
            Top             =   3780
            Width           =   2820
         End
         Begin VB.Line Line1 
            Index           =   41
            X1              =   2175
            X2              =   135
            Y1              =   4320
            Y2              =   4320
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu Selesai                       jam"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   51
            Left            =   135
            TabIndex        =   56
            Top             =   4095
            Width           =   2820
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   4740
         Index           =   2
         Left            =   7035
         TabIndex        =   4
         Top             =   1365
         Width           =   4755
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "C_1_selesai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   34
            Left            =   1020
            TabIndex        =   20
            Tag             =   "alkali"
            Top             =   1035
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "C_1_mulai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   35
            Left            =   1020
            TabIndex        =   19
            Tag             =   "alkali"
            Top             =   735
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "C_1_jumlah"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   36
            Left            =   1020
            TabIndex        =   18
            Tag             =   "alkali"
            Top             =   435
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "C_2_mulai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   25
            Left            =   990
            TabIndex        =   17
            Tag             =   "alkali"
            Top             =   1980
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "C_2_jumlah"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   26
            Left            =   990
            TabIndex        =   16
            Tag             =   "alkali"
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "C_2_selesai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   27
            Left            =   990
            TabIndex        =   15
            Tag             =   "alkali"
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "C_3_mulai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   28
            Left            =   990
            TabIndex        =   14
            Tag             =   "alkali"
            Top             =   3255
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "C_3_jumlah"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   29
            Left            =   990
            TabIndex        =   13
            Tag             =   "alkali"
            Top             =   2955
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "C_3_selesai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   30
            Left            =   990
            TabIndex        =   12
            Tag             =   "alkali"
            Top             =   3555
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "C_4_mulai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   31
            Left            =   3675
            TabIndex        =   11
            Tag             =   "alkali"
            Top             =   735
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "C_4_jumlah"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   32
            Left            =   3675
            TabIndex        =   10
            Tag             =   "alkali"
            Top             =   435
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "C_4_selesai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   33
            Left            =   3675
            TabIndex        =   9
            Tag             =   "alkali"
            Top             =   1035
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "C_5_mulai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   37
            Left            =   3675
            TabIndex        =   8
            Tag             =   "alkali"
            Top             =   1965
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "C_5_jumlah"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   38
            Left            =   3675
            TabIndex        =   7
            Tag             =   "alkali"
            Top             =   1665
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "C_5_selesai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   39
            Left            =   3675
            TabIndex        =   6
            Tag             =   "alkali"
            Top             =   2265
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ph"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   42
            Left            =   3675
            TabIndex        =   5
            Tag             =   "alkali"
            Top             =   2655
            Width           =   495
         End
         Begin VB.Line Line1 
            Index           =   35
            X1              =   1395
            X2              =   135
            Y1              =   1335
            Y2              =   1335
         End
         Begin VB.Line Line1 
            Index           =   36
            X1              =   1290
            X2              =   135
            Y1              =   1035
            Y2              =   1035
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Cuci 1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   41
            Left            =   150
            TabIndex        =   41
            Top             =   195
            Width           =   2715
         End
         Begin VB.Line Line1 
            Index           =   37
            X1              =   1500
            X2              =   135
            Y1              =   735
            Y2              =   735
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah               Liter"
            Height          =   255
            Index           =   29
            Left            =   135
            TabIndex        =   40
            Top             =   480
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Mulai"
            Height          =   255
            Index           =   30
            Left            =   120
            TabIndex        =   39
            Top             =   825
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Selesai"
            Height          =   255
            Index           =   32
            Left            =   150
            TabIndex        =   38
            Top             =   1140
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Cuci 2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   33
            Left            =   120
            TabIndex        =   37
            Top             =   1440
            Width           =   2715
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Cuci 3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   40
            Left            =   90
            TabIndex        =   30
            Top             =   2730
            Width           =   2715
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah                        Liter"
            Height          =   255
            Index           =   42
            Left            =   2295
            TabIndex        =   29
            Top             =   480
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Mulai"
            Height          =   255
            Index           =   43
            Left            =   2295
            TabIndex        =   28
            Top             =   795
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Selesai"
            Height          =   255
            Index           =   44
            Left            =   2280
            TabIndex        =   27
            Top             =   1095
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Cuci 4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   45
            Left            =   2295
            TabIndex        =   26
            Top             =   195
            Width           =   2715
         End
         Begin VB.Line Line1 
            Index           =   26
            X1              =   1320
            X2              =   90
            Y1              =   1980
            Y2              =   1980
         End
         Begin VB.Line Line1 
            Index           =   27
            X1              =   1320
            X2              =   105
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Line Line1 
            Index           =   28
            X1              =   1275
            X2              =   90
            Y1              =   2580
            Y2              =   2580
         End
         Begin VB.Line Line1 
            Index           =   29
            X1              =   1290
            X2              =   90
            Y1              =   3255
            Y2              =   3255
         End
         Begin VB.Line Line1 
            Index           =   30
            X1              =   1320
            X2              =   105
            Y1              =   3555
            Y2              =   3555
         End
         Begin VB.Line Line1 
            Index           =   31
            X1              =   1320
            X2              =   105
            Y1              =   3855
            Y2              =   3855
         End
         Begin VB.Line Line1 
            Index           =   32
            X1              =   3705
            X2              =   2265
            Y1              =   735
            Y2              =   735
         End
         Begin VB.Line Line1 
            Index           =   33
            X1              =   3720
            X2              =   2280
            Y1              =   1035
            Y2              =   1035
         End
         Begin VB.Line Line1 
            Index           =   34
            X1              =   3720
            X2              =   2280
            Y1              =   1335
            Y2              =   1335
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah                       Liter"
            Height          =   255
            Index           =   46
            Left            =   2310
            TabIndex        =   25
            Top             =   1725
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Mulai"
            Height          =   255
            Index           =   47
            Left            =   2295
            TabIndex        =   24
            Top             =   2025
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Selesai"
            Height          =   255
            Index           =   48
            Left            =   2280
            TabIndex        =   23
            Top             =   2310
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Cuci 5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   49
            Left            =   2295
            TabIndex        =   22
            Top             =   1425
            Width           =   2715
         End
         Begin VB.Line Line1 
            Index           =   38
            X1              =   3720
            X2              =   2280
            Y1              =   1965
            Y2              =   1965
         End
         Begin VB.Line Line1 
            Index           =   39
            X1              =   3720
            X2              =   2280
            Y1              =   2265
            Y2              =   2265
         End
         Begin VB.Line Line1 
            Index           =   40
            X1              =   3720
            X2              =   2280
            Y1              =   2565
            Y2              =   2565
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "PH"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   52
            Left            =   2310
            TabIndex        =   21
            Top             =   2745
            Width           =   360
         End
         Begin VB.Line Line1 
            Index           =   42
            X1              =   3735
            X2              =   2295
            Y1              =   2955
            Y2              =   2955
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Mulai"
            Height          =   255
            Index           =   35
            Left            =   105
            TabIndex        =   35
            Top             =   2055
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Selesai"
            Height          =   255
            Index           =   36
            Left            =   90
            TabIndex        =   34
            Top             =   2370
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah               Liter"
            Height          =   255
            Index           =   34
            Left            =   105
            TabIndex        =   36
            Top             =   1740
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah               Liter"
            Height          =   255
            Index           =   37
            Left            =   105
            TabIndex        =   33
            Top             =   3015
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Mulai"
            Height          =   255
            Index           =   38
            Left            =   105
            TabIndex        =   32
            Top             =   3345
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Selesai"
            Height          =   255
            Index           =   39
            Left            =   105
            TabIndex        =   31
            Top             =   3660
            Width           =   2865
         End
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "No_ekstrasi"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   0
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   3
         Tag             =   "alkali"
         Top             =   75
         Width           =   2160
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "desk"
         DataSource      =   "DDE"
         Height          =   705
         Index           =   40
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Tag             =   "alkali"
         Top             =   6480
         Width           =   3390
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "AT_tanggal"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   0
         Left            =   2130
         TabIndex        =   96
         Tag             =   "alkali"
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   51445763
         CurrentDate     =   39365
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   2220
         X2              =   180
         Y1              =   1275
         Y2              =   1275
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   2235
         X2              =   195
         Y1              =   975
         Y2              =   975
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2235
         X2              =   195
         Y1              =   675
         Y2              =   675
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2250
         X2              =   210
         Y1              =   375
         Y2              =   375
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tangki"
         Height          =   255
         Index           =   12
         Left            =   195
         TabIndex        =   101
         Top             =   1050
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         Height          =   255
         Index           =   2
         Left            =   195
         TabIndex        =   100
         Top             =   750
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   255
         Index           =   1
         Left            =   195
         TabIndex        =   99
         Top             =   435
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "No Ekstrasi"
         Height          =   255
         Index           =   0
         Left            =   195
         TabIndex        =   98
         Top             =   150
         Width           =   2055
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   50
         Left            =   90
         TabIndex        =   97
         Top             =   6225
         Width           =   2055
      End
   End
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   7380
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   1005
      BindFormTAG     =   "alkali"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmAlkali"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
Case tmbAddNew:
       txt(0).Text = IndexAuto
End Select
End Sub

Private Sub DDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
  
  Case tmbSave:
       DDE.IsChildMemberReady = True
       simpan
  Case tmbDelete:
       DDE.IsChildMemberReady = True
       Del
End Select
End Sub

Private Sub Form_Load()
With DDE
Set .BindForm = Me
    .BindFormTAG = "Alkali"
Set .ActiveConnection = CNN
    .PrepareQuery = "select * from Alkali_Treatmen"
End With
HiasFormManTell Picture2, Me
seting Me
End Sub
Function Del()
DDE.PrepareDelete = "delete  from Alkali_Treatmen where no_ekstrasi = '" + DDE.GetFieldByName("no_ekstrasi") + "'"
End Function

Private Function IndexAuto() As String
Dim Rc As New DBQuick
Dim TglSaiki As String
Dim Inom As Long
TglSaiki = Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2)
Rc.DBOpen "SELECT MAX(RIGHT(no_ekstrasi, 5)) AS MaxNom FROM [Alkali_treatmen] WHERE (GETDATE() = { fn NOW() })", CNN, lckLockReadOnly
With Rc
     If .DBRecordset.Recordcount <> 0 Then
        Inom = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + 1
     Else
        Inom = 1
     End If
     Select Case Len(Trim(Str(Inom)))
            Case 0: IndexAuto = "AT/" & TglSaiki & "-" & Trim(Str(Inom))
            Case 1: IndexAuto = "AT/" & TglSaiki & "-" & "0000" & Trim(Str(Inom))
            Case 2: IndexAuto = "AT/" & TglSaiki & "-" & "000" & Trim(Str(Inom))
            Case 3: IndexAuto = "AT/" & TglSaiki & "-" & "00" & Trim(Str(Inom))
            Case 4: IndexAuto = "AT/" & TglSaiki & "-" & "0" & Trim(Str(Inom))
     End Select
End With
End Function

Function simpan()

With DDE
   .PrepareAppend = "insert into alkali_treatmen (no_ekstrasi,AT_tanggal,Grup,tangki, no_stock, berat, jml_alkali_bekas,konsentrasi_alkali_bekas," & _
                                                  " jumlah_air,jml_alkali_baru, konsentrasi_alkali_baru,jml_alkali_akhir,konsentrasi_alkali_akhir, suhu_alkali_akhir," & _
                                                  " suhu_1, suhu_2, suhu_3, suhu_4, suhu_5, suhu_6, suhu_7, suhu_8, suhu_9, suhu_10, suhu_11, suhu_12," & _
                                                  " c_1_jumlah, c_1_mulai, c_1_selesai," & _
                                                  " c_2_jumlah, c_2_mulai, c_2_selesai," & _
                                                  " c_3_jumlah, c_3_mulai, c_3_selesai," & _
                                                  " c_4_jumlah, c_4_mulai, c_4_selesai," & _
                                                  " c_5_jumlah, c_5_mulai, c_5_selesai,desk,ph.waktu_selesai)" & _
                                                  " values('" & txt(0).Text & "','" & Format(DTPicker1(0).value, "yyyy-MM-dd") & "', " & _
                                                  " '" & DDE.GetFieldByName("Grup") & "', '" & DDE.GetFieldByName("Tangki") & "', " & _
                                                  " '" & DDE.GetFieldByName("no_stock") & "', '" & DDE.GetFieldByName("berat") & "', " & _
                                                  " '" & DDE.GetFieldByName("jml_alkali_bekas") & "', '" & DDE.GetFieldByName("konsentrasi_alkali_bekas") & "', '" & DDE.GetFieldByName("jumlah_air") & "', " & _
                                                  " '" & DDE.GetFieldByName("jml_alkali_baru") & "', '" & DDE.GetFieldByName("konsentrasi_alkali_baru") & "', " & _
                                                  " '" & DDE.GetFieldByName("jml_alkali_akhir") & "', '" & DDE.GetFieldByName("konsentrasi_alkali_akhir") & "', " & _
                                                  " '" & DDE.GetFieldByName("suhu_alkali_akhir") & "', '" & DDE.GetFieldByName("suhu_1") & "', " & _
                                                  " '" & DDE.GetFieldByName("suhu_2") & "', '" & DDE.GetFieldByName("suhu_3") & "', '" & DDE.GetFieldByName("suhu_4") & "', '" + DDE.GetFieldByName("suhu_5") & "', " & _
                                                  " '" & DDE.GetFieldByName("suhu_6") & "', '" & DDE.GetFieldByName("suhu_7") & "', '" & DDE.GetFieldByName("suhu_8") & "', '" + DDE.GetFieldByName("suhu_9") & "', " & _
                                                  " '" & DDE.GetFieldByName("suhu_10") & "', '" & DDE.GetFieldByName("suhu_11") & "', '" & DDE.GetFieldByName("suhu_12") & "', " & _
                                                  " '" & DDE.GetFieldByName("c_1_jumlah") & "','" & DDE.GetFieldByName("c_1_mulai") & "','" & DDE.GetFieldByName("c_1_selesai") & "', " & _
                                                  " '" & DDE.GetFieldByName("c_2_jumlah") & "','" & DDE.GetFieldByName("c_2_mulai") & "','" & DDE.GetFieldByName("c_2_selesai") & "', " & _
                                                  " '" & DDE.GetFieldByName("c_3_jumlah") & "','" & DDE.GetFieldByName("c_3_mulai") & "','" & DDE.GetFieldByName("c_3_selesai") & "', " & _
                                                  " '" & DDE.GetFieldByName("c_4_jumlah") & "','" & DDE.GetFieldByName("c_4_mulai") & "','" & DDE.GetFieldByName("c_4_selesai") & "', " & _
                                                  " '" & DDE.GetFieldByName("c_5_jumlah") & "','" & DDE.GetFieldByName("c_5_mulai") & "','" & DDE.GetFieldByName("c_5_selesai") & "','" & DDE.GetFieldByName("desk") & "','" & DDE.GetFieldByName("ph") & "','" & DDE.GetFieldByName("waktu_selesai") & "')"



.PrepareUpdate = " update alkali_treatmen set AT_tanggal = '" & Format(DTPicker1(0).value, "yyyy-MM-dd") & "', Grup = '" & DDE.GetFieldByName("Grup") & "', tangki = '" & DDE.GetFieldByName("Tangki") & "', " & _
                 " no_stock = '" & DDE.GetFieldByName("no_stock") & "', berat = '" & DDE.GetFieldByName("berat") & "', jml_alkali_bekas = '" & DDE.GetFieldByName("jml_alkali_bekas") & "', " & _
                 " konsentrasi_alkali_bekas = '" & DDE.GetFieldByName("konsentrasi_alkali_bekas") & "', jumlah_air= '" & DDE.GetFieldByName("jumlah_air") & "', jml_alkali_baru= '" & DDE.GetFieldByName("jml_alkali_baru") & "', " & _
                 " konsentrasi_alkali_baru= '" & DDE.GetFieldByName("konsentrasi_alkali_baru") & "', jml_alkali_akhir= '" & DDE.GetFieldByName("jml_alkali_akhir") & "', konsentrasi_alkali_akhir = '" & DDE.GetFieldByName("konsentrasi_alkali_akhir") & "', " & _
                 " suhu_alkali_akhir= '" & DDE.GetFieldByName("suhu_alkali_akhir") & "',  " & _
                 " suhu_1= '" & DDE.GetFieldByName("suhu_1") & "', suhu_2= '" & DDE.GetFieldByName("suhu_2") & "',suhu_3= '" & DDE.GetFieldByName("suhu_3") & "', " & _
                 " suhu_4= '" & DDE.GetFieldByName("suhu_4") & "', suhu_5= '" & DDE.GetFieldByName("suhu_5") & "', suhu_6= '" & DDE.GetFieldByName("suhu_6") & "', " & _
                 " suhu_7= '" & DDE.GetFieldByName("suhu_7") & "', suhu_8= '" & DDE.GetFieldByName("suhu_8") & "', suhu_9= '" & DDE.GetFieldByName("suhu_9") & "', " & _
                 " suhu_10= '" & DDE.GetFieldByName("suhu_10") & "', suhu_11= '" & DDE.GetFieldByName("suhu_11") & "', suhu_12= '" & DDE.GetFieldByName("suhu_12") & "', " & _
                 " c_1_jumlah= '" & DDE.GetFieldByName("c_1_jumlah") & "', c_1_mulai= '" & DDE.GetFieldByName("c_1_mulai") & "', c_1_selesai= '" & DDE.GetFieldByName("c_1_selesai") & "', " & _
                 " c_2_jumlah= '" & DDE.GetFieldByName("c_2_jumlah") & "', c_2_mulai= '" & DDE.GetFieldByName("c_2_mulai") & "', c_2_selesai= '" & DDE.GetFieldByName("c_2_selesai") & "', " & _
                 " c_3_jumlah= '" & DDE.GetFieldByName("c_3_jumlah") & "', c_3_mulai= '" & DDE.GetFieldByName("c_3_mulai") & "',c_3_selesai= '" & DDE.GetFieldByName("c_3_selesai") & "', " & _
                 " c_4_jumlah= '" & DDE.GetFieldByName("c_4_jumlah") & "', c_4_mulai= '" & DDE.GetFieldByName("c_4_mulai") & "', c_4_selesai= '" & DDE.GetFieldByName("c_4_selesai") & "', " & _
                 " c_5_jumlah= '" & DDE.GetFieldByName("c_5_jumlah") & "', c_5_mulai= '" & DDE.GetFieldByName("c_5_mulai") & "', c_5_selesai= '" & DDE.GetFieldByName("c_5_selesai") & "', " & _
                 " desk= '" & DDE.GetFieldByName("desk") & "', ph = '" & DDE.GetFieldByName("ph") & "', waktu_selesai = '" & DDE.GetFieldByName("waktu_selesai") & "' where no_ekstrasi ='" & .GetFieldByName("no_ekstrasi") & "'"



End With
End Function

Private Sub txt_GotFocus(Index As Integer)
txt(Index).BackColor = &H79BCFF
End Sub

Private Sub txt_LostFocus(Index As Integer)
txt(Index).BackColor = vbWhite
End Sub
