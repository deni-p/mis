VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D7BB8F75-AC9E-4E80-A526-70EA20ACFD16}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmPAcidTreatment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acid Treatment"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9825
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   9825
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      ForeColor       =   &H80000008&
      Height          =   7110
      Left            =   0
      ScaleHeight     =   7080
      ScaleWidth      =   9810
      TabIndex        =   1
      Top             =   0
      Width           =   9840
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   4260
         Index           =   2
         Left            =   4140
         TabIndex        =   29
         Top             =   1590
         Width           =   5295
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_1_jumlah"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   36
            Left            =   1515
            TabIndex        =   48
            Tag             =   "ACID"
            Top             =   495
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_1_mulai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   35
            Left            =   1515
            TabIndex        =   47
            Tag             =   "ACID"
            Top             =   795
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_1_selesai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   34
            Left            =   1515
            TabIndex        =   46
            Tag             =   "ACID"
            Top             =   1095
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_3_selesai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   15
            Left            =   1485
            TabIndex        =   45
            Tag             =   "ACID"
            Top             =   2325
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_3_mulai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   16
            Left            =   1485
            TabIndex        =   44
            Tag             =   "ACID"
            Top             =   2025
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_3_jumlah"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   17
            Left            =   1485
            TabIndex        =   43
            Tag             =   "ACID"
            Top             =   1725
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_5_selesai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   18
            Left            =   1485
            TabIndex        =   42
            Tag             =   "ACID"
            Top             =   3540
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_5_mulai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   19
            Left            =   1485
            TabIndex        =   41
            Tag             =   "ACID"
            Top             =   3240
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_5_jumlah"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   20
            Left            =   1485
            TabIndex        =   40
            Tag             =   "ACID"
            Top             =   2940
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_2_selesai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   21
            Left            =   4305
            TabIndex        =   39
            Tag             =   "ACID"
            Top             =   1110
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_2_mulai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   22
            Left            =   4305
            TabIndex        =   38
            Tag             =   "ACID"
            Top             =   810
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_2_jumlah"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   23
            Left            =   4305
            TabIndex        =   37
            Tag             =   "ACID"
            Top             =   510
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_4_selesai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   24
            Left            =   4305
            TabIndex        =   36
            Tag             =   "ACID"
            Top             =   2310
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_4_mulai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   25
            Left            =   4305
            TabIndex        =   35
            Tag             =   "ACID"
            Top             =   2010
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_4_jumlah"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   26
            Left            =   4305
            TabIndex        =   34
            Tag             =   "ACID"
            Top             =   1710
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_6_selesai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   4
            Left            =   4275
            TabIndex        =   33
            Tag             =   "ACID"
            Top             =   3510
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_6_mulai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   5
            Left            =   4275
            TabIndex        =   32
            Tag             =   "ACID"
            Top             =   3210
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_6_jumlah"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   6
            Left            =   4275
            TabIndex        =   31
            Tag             =   "ACID"
            Top             =   2910
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ph"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   27
            Left            =   4275
            TabIndex        =   30
            Tag             =   "ACID"
            Top             =   3810
            Width           =   495
         End
         Begin VB.Line Line1 
            Index           =   40
            X1              =   1620
            X2              =   180
            Y1              =   3840
            Y2              =   3840
         End
         Begin VB.Line Line1 
            Index           =   39
            X1              =   1620
            X2              =   180
            Y1              =   3540
            Y2              =   3540
         End
         Begin VB.Line Line1 
            Index           =   38
            X1              =   1620
            X2              =   180
            Y1              =   3240
            Y2              =   3240
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
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   49
            Left            =   195
            TabIndex        =   73
            Top             =   2685
            Width           =   2715
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Selesai"
            Height          =   255
            Index           =   48
            Left            =   180
            TabIndex        =   72
            Top             =   3585
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Mulai"
            Height          =   255
            Index           =   47
            Left            =   195
            TabIndex        =   71
            Top             =   3300
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah                                Liter"
            Height          =   255
            Index           =   46
            Left            =   210
            TabIndex        =   70
            Top             =   3000
            Width           =   2865
         End
         Begin VB.Line Line1 
            Index           =   34
            X1              =   4305
            X2              =   2865
            Y1              =   2610
            Y2              =   2610
         End
         Begin VB.Line Line1 
            Index           =   33
            X1              =   4305
            X2              =   2865
            Y1              =   2310
            Y2              =   2310
         End
         Begin VB.Line Line1 
            Index           =   32
            X1              =   4290
            X2              =   2850
            Y1              =   2010
            Y2              =   2010
         End
         Begin VB.Line Line1 
            Index           =   31
            X1              =   1605
            X2              =   165
            Y1              =   2625
            Y2              =   2625
         End
         Begin VB.Line Line1 
            Index           =   30
            X1              =   1590
            X2              =   150
            Y1              =   2325
            Y2              =   2325
         End
         Begin VB.Line Line1 
            Index           =   29
            X1              =   1575
            X2              =   135
            Y1              =   2025
            Y2              =   2025
         End
         Begin VB.Line Line1 
            Index           =   28
            X1              =   4305
            X2              =   2865
            Y1              =   1410
            Y2              =   1410
         End
         Begin VB.Line Line1 
            Index           =   27
            X1              =   4305
            X2              =   2865
            Y1              =   1110
            Y2              =   1110
         End
         Begin VB.Line Line1 
            Index           =   26
            X1              =   4305
            X2              =   2865
            Y1              =   810
            Y2              =   810
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
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   45
            Left            =   2880
            TabIndex        =   69
            Top             =   1485
            Width           =   2715
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Selesai"
            Height          =   255
            Index           =   44
            Left            =   2865
            TabIndex        =   68
            Top             =   2385
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Mulai"
            Height          =   255
            Index           =   43
            Left            =   2880
            TabIndex        =   67
            Top             =   2070
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah                                Liter"
            Height          =   255
            Index           =   42
            Left            =   2880
            TabIndex        =   66
            Top             =   1755
            Width           =   2865
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
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   40
            Left            =   165
            TabIndex        =   65
            Top             =   1485
            Width           =   2715
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Selesai"
            Height          =   255
            Index           =   39
            Left            =   180
            TabIndex        =   64
            Top             =   2415
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Mulai"
            Height          =   255
            Index           =   38
            Left            =   180
            TabIndex        =   63
            Top             =   2100
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah                                Liter"
            Height          =   255
            Index           =   37
            Left            =   165
            TabIndex        =   62
            Top             =   1785
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Selesai"
            Height          =   255
            Index           =   36
            Left            =   2865
            TabIndex        =   61
            Top             =   1200
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Mulai"
            Height          =   255
            Index           =   35
            Left            =   2880
            TabIndex        =   60
            Top             =   900
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah                                Liter"
            Height          =   255
            Index           =   34
            Left            =   2880
            TabIndex        =   59
            Top             =   570
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
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   33
            Left            =   2895
            TabIndex        =   58
            Top             =   270
            Width           =   2715
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Selesai"
            Height          =   255
            Index           =   32
            Left            =   150
            TabIndex        =   57
            Top             =   1200
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Mulai"
            Height          =   255
            Index           =   30
            Left            =   135
            TabIndex        =   56
            Top             =   885
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah                                Liter"
            Height          =   255
            Index           =   29
            Left            =   135
            TabIndex        =   55
            Top             =   540
            Width           =   2865
         End
         Begin VB.Line Line1 
            Index           =   37
            X1              =   1500
            X2              =   135
            Y1              =   795
            Y2              =   795
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
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   41
            Left            =   150
            TabIndex        =   54
            Top             =   255
            Width           =   2715
         End
         Begin VB.Line Line1 
            Index           =   36
            X1              =   1815
            X2              =   135
            Y1              =   1095
            Y2              =   1095
         End
         Begin VB.Line Line1 
            Index           =   35
            X1              =   1830
            X2              =   135
            Y1              =   1395
            Y2              =   1395
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Selesai"
            Height          =   255
            Index           =   6
            Left            =   2850
            TabIndex        =   53
            Top             =   3585
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Mulai"
            Height          =   255
            Index           =   7
            Left            =   2865
            TabIndex        =   52
            Top             =   3300
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah                                Liter"
            Height          =   255
            Index           =   17
            Left            =   2865
            TabIndex        =   51
            Top             =   2970
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Cuci 6"
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
            Left            =   2865
            TabIndex        =   50
            Top             =   2715
            Width           =   2715
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   4290
            X2              =   2850
            Y1              =   3210
            Y2              =   3210
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   4290
            X2              =   2850
            Y1              =   3510
            Y2              =   3510
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   4260
            X2              =   2820
            Y1              =   3810
            Y2              =   3810
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
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   20
            Left            =   2970
            TabIndex        =   49
            Top             =   3900
            Width           =   345
         End
         Begin VB.Line Line1 
            Index           =   16
            X1              =   4530
            X2              =   2820
            Y1              =   4110
            Y2              =   4110
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   4275
         Index           =   0
         Left            =   315
         TabIndex        =   7
         Top             =   1590
         Width           =   3720
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "suhu_acid_akhir"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   12
            Left            =   2085
            TabIndex        =   16
            Tag             =   "ACID"
            Top             =   2640
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "konsentrasi_acid_akhir"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   11
            Left            =   2085
            TabIndex        =   15
            Tag             =   "ACID"
            Top             =   2340
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "jml_acid_akhir"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   10
            Left            =   2085
            TabIndex        =   14
            Tag             =   "ACID"
            Top             =   2040
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "konsentrasi_acid"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   9
            Left            =   2085
            TabIndex        =   13
            Tag             =   "ACID"
            Top             =   1395
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "JML_ACID"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   8
            Left            =   2085
            TabIndex        =   12
            Tag             =   "ACID"
            Top             =   1095
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "jumlah_air"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   7
            Left            =   2085
            TabIndex        =   11
            Tag             =   "ACID"
            Top             =   465
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "waktu_mulai_treatmen"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   3
            Left            =   2085
            TabIndex        =   10
            Tag             =   "ACID"
            Top             =   3045
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "suhu"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   13
            Left            =   2085
            TabIndex        =   9
            Tag             =   "ACID"
            Top             =   3345
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "waktu_selesai_treatmen"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   14
            Left            =   2085
            TabIndex        =   8
            Tag             =   "ACID"
            Top             =   3645
            Width           =   495
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Konsentrasi                                     %"
            Height          =   255
            Index           =   14
            Left            =   165
            TabIndex        =   28
            Top             =   1485
            Width           =   2820
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Konsentrasi                                     %"
            Height          =   255
            Index           =   16
            Left            =   165
            TabIndex        =   27
            Top             =   2415
            Width           =   2820
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah                                            Liter"
            Height          =   255
            Index           =   13
            Left            =   180
            TabIndex        =   26
            Top             =   2130
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah                                            Kg"
            Height          =   255
            Index           =   11
            Left            =   180
            TabIndex        =   25
            Top             =   1170
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah                                            Liter"
            Height          =   255
            Index           =   8
            Left            =   150
            TabIndex        =   24
            Top             =   555
            Width           =   2865
         End
         Begin VB.Line Line1 
            Index           =   13
            X1              =   2190
            X2              =   150
            Y1              =   2940
            Y2              =   2940
         End
         Begin VB.Line Line1 
            Index           =   12
            X1              =   2190
            X2              =   150
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Line Line1 
            Index           =   11
            X1              =   2190
            X2              =   150
            Y1              =   2340
            Y2              =   2340
         End
         Begin VB.Line Line1 
            Index           =   10
            X1              =   2190
            X2              =   150
            Y1              =   1695
            Y2              =   1695
         End
         Begin VB.Line Line1 
            Index           =   9
            X1              =   2205
            X2              =   165
            Y1              =   1395
            Y2              =   1395
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   2175
            X2              =   135
            Y1              =   765
            Y2              =   765
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Suhu                                                C"
            Height          =   255
            Index           =   18
            Left            =   165
            TabIndex        =   23
            Top             =   2700
            Width           =   2730
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Larutan ACID Akhir"
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
            Index           =   15
            Left            =   165
            TabIndex        =   22
            Top             =   1830
            Width           =   2055
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "ACID"
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
            Index           =   10
            Left            =   165
            TabIndex        =   21
            Top             =   870
            Width           =   2055
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
            Left            =   165
            TabIndex        =   20
            Top             =   240
            Width           =   2055
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   2190
            X2              =   150
            Y1              =   3345
            Y2              =   3345
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu Mulai Treatment"
            Height          =   255
            Index           =   3
            Left            =   150
            TabIndex        =   19
            Top             =   3090
            Width           =   1875
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Suhu pada 20 menit                         C"
            Height          =   255
            Index           =   4
            Left            =   165
            TabIndex        =   18
            Top             =   3435
            Width           =   2910
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu Selesai Treatment"
            Height          =   255
            Index           =   5
            Left            =   150
            TabIndex        =   17
            Top             =   3735
            Width           =   1875
         End
         Begin VB.Line Line1 
            Index           =   14
            X1              =   2190
            X2              =   150
            Y1              =   3645
            Y2              =   3645
         End
         Begin VB.Line Line1 
            Index           =   15
            X1              =   2190
            X2              =   150
            Y1              =   3945
            Y2              =   3945
         End
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "grup"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   6
         Tag             =   "ACID"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "no_ekstrasi"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   2295
         Locked          =   -1  'True
         TabIndex        =   5
         Tag             =   "ACID"
         Top             =   120
         Width           =   1830
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "tangki"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   2
         Left            =   2280
         TabIndex        =   4
         Tag             =   "ACID"
         Top             =   1020
         Width           =   1695
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4140
         MaskColor       =   &H000000C0&
         Picture         =   "frmAcidTreatment.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "SPPH"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "desk_acid"
         DataSource      =   "DDE"
         Height          =   840
         Index           =   40
         Left            =   345
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Tag             =   "ACID"
         Top             =   6165
         Width           =   3390
      End
      Begin MSComCtl2.DTPicker tanggal 
         DataField       =   "tanggal_acid"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   0
         Left            =   2280
         TabIndex        =   74
         Tag             =   "ACID"
         Top             =   405
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   61538307
         CurrentDate     =   39365
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "No Ekstrasi"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   79
         Top             =   165
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   78
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   77
         Top             =   795
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tangki"
         Height          =   255
         Index           =   12
         Left            =   360
         TabIndex        =   76
         Top             =   1095
         Width           =   2055
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2400
         X2              =   360
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2400
         X2              =   360
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   2400
         X2              =   360
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   2385
         X2              =   345
         Y1              =   1320
         Y2              =   1320
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
         Index           =   50
         Left            =   330
         TabIndex        =   75
         Top             =   5925
         Width           =   2055
      End
   End
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   7095
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   1005
      BindFormTAG     =   "TTRL"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmPAcidTreatment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents Bleacing  As frmCaller
Attribute Bleacing.VB_VarHelpID = -1

Private WithEvents mAlkali  As frmCaller
Attribute mAlkali.VB_VarHelpID = -1
Dim rsbleacing As New DBQuick
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

        Case tmbAddNew:
            CmdLink.Enabled = True
            txt(0).Enabled = False
    End Select

End Sub

Private Sub Form_Load()

    With DDE
        Set .BindForm = Me
        .BindFormTAG = "ACID"
        Set .ActiveConnection = CNN
        .PrepareQuery = "select * from ACID_TREATMEN"
    End With

    'HiasForm Picture1, Me
    HiasFormManTell Picture2, Me
    seting Me
End Sub

Function Del()
    DDE.PrepareDelete = "delete  from ACID_TREATMEN where no_ekstrasi = '" + DDE.GetFieldByName("no_ekstrasi") + "'"
End Function

Private Sub DDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb

        Case tmbSave:
            DDE.IsChildMemberReady = True
            simpan

        Case tmbDelete:
            Del
    End Select

End Sub

Function simpan()

    With DDE
        .PrepareAppend = "insert into ACID_TREATMEN (no_ekstrasi,tanggal_acid,Grup,tangki," & _
           " jumlah_air,jml_Acid, konsentrasi_Acid,jml_acid_akhir,konsentrasi_acid_akhir, suhu_acid_akhir," & _
           " waktu_mulai_treatmen,suhu,waktu_selesai_treatmen," & _
           " c_1_jumlah, c_1_mulai, c_1_selesai," & _
           " c_2_jumlah, c_2_mulai, c_2_selesai," & _
           " c_3_jumlah, c_3_mulai, c_3_selesai," & _
           " c_4_jumlah, c_4_mulai, c_4_selesai," & _
           " c_5_jumlah, c_5_mulai, c_5_selesai," & _
           " c_6_jumlah, c_6_mulai, c_6_selesai,desk_acid,ph) values" & _
           "('" + DDE.GetFieldByName("no_ekstrasi") + "','" + Format(tanggal(0).value, "yyyy-MM-dd") + "', " & _
           " '" + DDE.GetFieldByName("Grup") + "', '" + DDE.GetFieldByName("Tangki") + "', " & _
           " '" + DDE.GetFieldByName("jumlah_air") + "', " & _
           " '" + DDE.GetFieldByName("jml_acid") + "', '" + DDE.GetFieldByName("konsentrasi_acid") + "'," & _
           " '" + DDE.GetFieldByName("jml_acid_akhir") + "','" + DDE.GetFieldByName("konsentrasi_acid_akhir") + "', '" + DDE.GetFieldByName("suhu_acid_akhir") + "', " & _
           " '" + DDE.GetFieldByName("waktu_mulai_treatmen") + "', '" + DDE.GetFieldByName("suhu") + "', " & _
           " '" + DDE.GetFieldByName("waktu_selesai_treatmen") + "', " & _
           " '" + DDE.GetFieldByName("c_1_jumlah") + "','" + DDE.GetFieldByName("c_1_mulai") + "','" + DDE.GetFieldByName("c_1_selesai") + "', " & _
           " '" + DDE.GetFieldByName("c_2_jumlah") + "','" + DDE.GetFieldByName("c_2_mulai") + "','" + DDE.GetFieldByName("c_2_selesai") + "', " & _
           " '" + DDE.GetFieldByName("c_3_jumlah") + "','" + DDE.GetFieldByName("c_3_mulai") + "','" + DDE.GetFieldByName("c_3_selesai") + "', " & _
           " '" + DDE.GetFieldByName("c_4_jumlah") + "','" + DDE.GetFieldByName("c_4_mulai") + "','" + DDE.GetFieldByName("c_4_selesai") + "','" + DDE.GetFieldByName("c_5_jumlah") + "','" + DDE.GetFieldByName("c_5_mulai") + "','" + DDE.GetFieldByName("c_5_selesai") + "', " & _
           " '" + DDE.GetFieldByName("c_6_jumlah") + "','" + DDE.GetFieldByName("c_6_mulai") + "','" + DDE.GetFieldByName("c_6_selesai") + "','" + DDE.GetFieldByName("desk_acid") + "','" + DDE.GetFieldByName("ph") + "')"
    
        .PrepareUpdate = "update ACID_TREATMEN set tanggal_acid = '" & Format(tanggal(0).value, "yyyy-MM-dd") & "',  " & _
           " Grup = '" & DDE.GetFieldByName("Grup") & "', tangki = '" & DDE.GetFieldByName("Tangki") & "', jumlah_air= '" & DDE.GetFieldByName("jumlah_air") & "' , " & _
           " jml_Acid= '" & DDE.GetFieldByName("jml_acid") & "', konsentrasi_Acid= '" & DDE.GetFieldByName("konsentrasi_acid") & "', jml_acid_akhir= '" & DDE.GetFieldByName("jml_acid_akhir") & "', " & _
           " konsentrasi_acid_akhir = '" & DDE.GetFieldByName("konsentrasi_acid_akhir") & "', suhu_acid_akhir= '" & DDE.GetFieldByName("suhu_acid_akhir") & "', waktu_mulai_treatmen= '" & DDE.GetFieldByName("waktu_mulai_treatmen") & "', " & _
           " suhu= '" & DDE.GetFieldByName("suhu") & "', waktu_selesai_treatmen = '" & DDE.GetFieldByName("waktu_selesai_treatmen") & "', " & _
           " c_1_jumlah = '" & DDE.GetFieldByName("c_1_jumlah") & "', c_1_mulai= '" & DDE.GetFieldByName("c_1_mulai") & "', c_1_selesai= '" & DDE.GetFieldByName("c_1_selesai") & "', " & _
           " c_2_jumlah= '" & DDE.GetFieldByName("c_2_jumlah") & "', c_2_mulai= '" & DDE.GetFieldByName("c_2_mulai") & "', c_2_selesai= '" & DDE.GetFieldByName("c_2_selesai") & "', " & _
           " c_3_jumlah= '" & DDE.GetFieldByName("c_3_jumlah") & "', c_3_mulai= '" & DDE.GetFieldByName("c_3_mulai") & "', c_3_selesai= '" & DDE.GetFieldByName("c_3_selesai") & "', " & _
           " c_4_jumlah= '" & DDE.GetFieldByName("c_4_jumlah") & "', c_4_mulai= '" & DDE.GetFieldByName("c_4_mulai") & "', c_4_selesai= '" & DDE.GetFieldByName("c_4_selesai") & "', " & _
           " c_5_jumlah= '" & DDE.GetFieldByName("c_5_jumlah") & "', c_5_mulai= '" & DDE.GetFieldByName("c_5_mulai") & "', c_5_selesai= '" & DDE.GetFieldByName("c_5_selesai") & "', " & _
           " c_6_jumlah= '" & DDE.GetFieldByName("c_6_jumlah") & "', c_6_mulai= '" & DDE.GetFieldByName("c_6_mulai") & "', c_6_selesai= '" & DDE.GetFieldByName("c_6_selesai") & "', " & _
           " desk_acid= '" & DDE.GetFieldByName("desk_acid") & "', ph = '" & DDE.GetFieldByName("ph") & "' where no_ekstrasi ='" & .GetFieldByName("no_ekstrasi") & "'"

    End With

End Function

Private Sub mAlkali_RowColChange(ByVal TagForm As String, _
                                 ByVal pRecordset As ADODB.Recordset)
    DDE.GetFieldByName("no_ekstrasi") = rsalkali.DBRecordset.Fields("no_ekstrasi")
End Sub

Private Sub txt_GotFocus(Index As Integer)
    txt(Index).BackColor = &H79BCFF
End Sub

Private Sub txt_LostFocus(Index As Integer)
    txt(Index).BackColor = vbWhite
End Sub
