VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D7BB8F75-AC9E-4E80-A526-70EA20ACFD16}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmPbleacing 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bleaching Treatment"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10500
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      ForeColor       =   &H80000008&
      Height          =   6915
      Left            =   0
      ScaleHeight     =   6885
      ScaleWidth      =   10515
      TabIndex        =   1
      Top             =   0
      Width           =   10545
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "tangki"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   2
         Left            =   2280
         TabIndex        =   74
         Tag             =   "bleaching"
         Top             =   1020
         Width           =   1695
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "no_ekstrasi"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   2280
         TabIndex        =   73
         Tag             =   "bleaching"
         Top             =   120
         Width           =   1845
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "grup"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   72
         Tag             =   "bleaching"
         Top             =   720
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   4260
         Index           =   0
         Left            =   315
         TabIndex        =   50
         Top             =   1425
         Width           =   3720
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "waktu_selesai_treatmen"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   14
            Left            =   2085
            TabIndex        =   59
            Tag             =   "bleaching"
            Top             =   3645
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "suhu"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   13
            Left            =   2085
            TabIndex        =   58
            Tag             =   "bleaching"
            Top             =   3345
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "waktu_mulai_treatmen"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   3
            Left            =   2085
            TabIndex        =   57
            Tag             =   "bleaching"
            Top             =   3045
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "jumlah_air"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   7
            Left            =   2085
            TabIndex        =   56
            Tag             =   "bleaching"
            Top             =   465
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "JML_BLEACHING"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   8
            Left            =   2085
            TabIndex        =   55
            Tag             =   "bleaching"
            Top             =   1095
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "konsentrasi_BLEACHING"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   9
            Left            =   2085
            TabIndex        =   54
            Tag             =   "bleaching"
            Top             =   1395
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "jml_BLEACHING_akhir"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   10
            Left            =   2085
            TabIndex        =   53
            Tag             =   "bleaching"
            Top             =   2040
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "konsentrasi_BLEACHING_akhir"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   11
            Left            =   2085
            TabIndex        =   52
            Tag             =   "bleaching"
            Top             =   2340
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "suhu_BLEACHING_akhir"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   12
            Left            =   2085
            TabIndex        =   51
            Tag             =   "bleaching"
            Top             =   2640
            Width           =   495
         End
         Begin VB.Line Line1 
            Index           =   15
            X1              =   2190
            X2              =   150
            Y1              =   3945
            Y2              =   3945
         End
         Begin VB.Line Line1 
            Index           =   14
            X1              =   2190
            X2              =   150
            Y1              =   3645
            Y2              =   3645
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu Selesai Treatment"
            Height          =   255
            Index           =   5
            Left            =   150
            TabIndex        =   71
            Top             =   3735
            Width           =   1875
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Suhu pada 20 menit                         C"
            Height          =   255
            Index           =   4
            Left            =   165
            TabIndex        =   70
            Top             =   3435
            Width           =   2910
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu Mulai Treatment"
            Height          =   255
            Index           =   3
            Left            =   150
            TabIndex        =   69
            Top             =   3090
            Width           =   1875
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
            TabIndex        =   68
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Bleaching Agent"
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
            TabIndex        =   67
            Top             =   870
            Width           =   2055
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Larutan Bleaching Akhir"
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
            TabIndex        =   66
            Top             =   1830
            Width           =   2055
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Suhu                                                C"
            Height          =   255
            Index           =   18
            Left            =   165
            TabIndex        =   65
            Top             =   2700
            Width           =   2730
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   2175
            X2              =   135
            Y1              =   765
            Y2              =   765
         End
         Begin VB.Line Line1 
            Index           =   9
            X1              =   2205
            X2              =   165
            Y1              =   1395
            Y2              =   1395
         End
         Begin VB.Line Line1 
            Index           =   10
            X1              =   2190
            X2              =   150
            Y1              =   1695
            Y2              =   1695
         End
         Begin VB.Line Line1 
            Index           =   11
            X1              =   2190
            X2              =   150
            Y1              =   2340
            Y2              =   2340
         End
         Begin VB.Line Line1 
            Index           =   12
            X1              =   2190
            X2              =   150
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Line Line1 
            Index           =   13
            X1              =   2190
            X2              =   150
            Y1              =   2940
            Y2              =   2940
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah                                            Liter"
            Height          =   255
            Index           =   8
            Left            =   150
            TabIndex        =   64
            Top             =   555
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah                                            Kg"
            Height          =   255
            Index           =   11
            Left            =   180
            TabIndex        =   63
            Top             =   1170
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah                                            Liter"
            Height          =   255
            Index           =   13
            Left            =   180
            TabIndex        =   62
            Top             =   2130
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Konsentrasi                                     %"
            Height          =   255
            Index           =   16
            Left            =   165
            TabIndex        =   61
            Top             =   2415
            Width           =   2820
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Konsentrasi                                     %"
            Height          =   255
            Index           =   14
            Left            =   165
            TabIndex        =   60
            Top             =   1485
            Width           =   2820
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   4245
         Index           =   2
         Left            =   4140
         TabIndex        =   4
         Top             =   1440
         Width           =   6090
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "ph"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   27
            Left            =   3495
            TabIndex        =   24
            Tag             =   "bleaching"
            Top             =   3750
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_6_jumlah"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   6
            Left            =   3495
            TabIndex        =   23
            Tag             =   "bleaching"
            Top             =   2850
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_6_mulai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   5
            Left            =   3495
            TabIndex        =   22
            Tag             =   "bleaching"
            Top             =   3150
            Width           =   1860
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_6_selesai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   4
            Left            =   3495
            TabIndex        =   21
            Tag             =   "bleaching"
            Top             =   3450
            Width           =   1860
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_4_jumlah"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   26
            Left            =   3495
            TabIndex        =   20
            Tag             =   "bleaching"
            Top             =   1665
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_4_mulai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   25
            Left            =   3495
            TabIndex        =   19
            Tag             =   "bleaching"
            Top             =   1965
            Width           =   1860
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_4_selesai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   24
            Left            =   3495
            TabIndex        =   18
            Tag             =   "bleaching"
            Top             =   2265
            Width           =   1860
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_2_jumlah"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   23
            Left            =   3495
            TabIndex        =   17
            Tag             =   "bleaching"
            Top             =   465
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_2_mulai"
            DataSource      =   "DDE"
            Height          =   300
            Index           =   22
            Left            =   3495
            TabIndex        =   16
            Tag             =   "bleaching"
            Top             =   765
            Width           =   1860
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_2_selesai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   21
            Left            =   3495
            TabIndex        =   15
            Tag             =   "bleaching"
            Top             =   1050
            Width           =   1860
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_5_jumlah"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   20
            Left            =   945
            TabIndex        =   14
            Tag             =   "bleaching"
            Top             =   2880
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_5_mulai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   19
            Left            =   945
            TabIndex        =   13
            Tag             =   "bleaching"
            Top             =   3180
            Width           =   1875
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_5_selesai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   18
            Left            =   945
            TabIndex        =   12
            Tag             =   "bleaching"
            Top             =   3480
            Width           =   1875
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_3_jumlah"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   17
            Left            =   945
            TabIndex        =   11
            Tag             =   "bleaching"
            Top             =   1665
            Width           =   495
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_3_mulai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   16
            Left            =   945
            TabIndex        =   10
            Tag             =   "bleaching"
            Top             =   1965
            Width           =   1875
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_3_selesai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   15
            Left            =   945
            TabIndex        =   9
            Tag             =   "bleaching"
            Top             =   2265
            Width           =   1875
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_1_selesai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   34
            Left            =   945
            TabIndex        =   8
            Tag             =   "bleaching"
            Top             =   1035
            Width           =   1875
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_1_mulai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   35
            Left            =   945
            TabIndex        =   7
            Tag             =   "bleaching"
            Top             =   735
            Width           =   1875
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "c_1_jumlah"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   36
            Left            =   945
            TabIndex        =   6
            Tag             =   "bleaching"
            Top             =   435
            Width           =   495
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Left            =   3495
            TabIndex        =   5
            Top             =   1260
            Visible         =   0   'False
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "hh:mm:ss"
            Format          =   58458114
            CurrentDate     =   39443
         End
         Begin VB.Line Line1 
            Index           =   16
            X1              =   3900
            X2              =   2850
            Y1              =   4050
            Y2              =   4050
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
            Left            =   2910
            TabIndex        =   49
            Top             =   3825
            Width           =   345
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   4290
            X2              =   2850
            Y1              =   3750
            Y2              =   3750
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   4320
            X2              =   2880
            Y1              =   3450
            Y2              =   3450
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   4320
            X2              =   2880
            Y1              =   3150
            Y2              =   3150
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
            Left            =   2895
            TabIndex        =   48
            Top             =   2655
            Width           =   2715
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah              Liter"
            Height          =   255
            Index           =   17
            Left            =   2895
            TabIndex        =   47
            Top             =   2910
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Mulai"
            Height          =   255
            Index           =   7
            Left            =   2895
            TabIndex        =   46
            Top             =   3240
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Selesai"
            Height          =   255
            Index           =   6
            Left            =   2880
            TabIndex        =   45
            Top             =   3525
            Width           =   2865
         End
         Begin VB.Line Line1 
            Index           =   35
            X1              =   1830
            X2              =   135
            Y1              =   1335
            Y2              =   1335
         End
         Begin VB.Line Line1 
            Index           =   36
            X1              =   1815
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
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   41
            Left            =   150
            TabIndex        =   44
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
            Caption         =   "Jumlah                   Liter"
            Height          =   255
            Index           =   29
            Left            =   135
            TabIndex        =   43
            Top             =   480
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Mulai"
            Height          =   255
            Index           =   30
            Left            =   135
            TabIndex        =   42
            Top             =   825
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Selesai"
            Height          =   255
            Index           =   32
            Left            =   150
            TabIndex        =   41
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
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   33
            Left            =   2895
            TabIndex        =   40
            Top             =   210
            Width           =   2715
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah              Liter"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   34
            Left            =   2880
            TabIndex        =   39
            Top             =   510
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Mulai"
            Height          =   255
            Index           =   35
            Left            =   2880
            TabIndex        =   38
            Top             =   840
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Selesai"
            Height          =   255
            Index           =   36
            Left            =   2865
            TabIndex        =   37
            Top             =   1140
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah                 Liter"
            Height          =   255
            Index           =   37
            Left            =   195
            TabIndex        =   36
            Top             =   1725
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Mulai"
            Height          =   255
            Index           =   38
            Left            =   210
            TabIndex        =   35
            Top             =   2040
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Selesai"
            Height          =   255
            Index           =   39
            Left            =   210
            TabIndex        =   34
            Top             =   2355
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
            Left            =   195
            TabIndex        =   33
            Top             =   1425
            Width           =   2715
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah              Liter"
            Height          =   255
            Index           =   42
            Left            =   2880
            TabIndex        =   32
            Top             =   1710
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Mulai"
            Height          =   255
            Index           =   43
            Left            =   2880
            TabIndex        =   31
            Top             =   2025
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Selesai"
            Height          =   255
            Index           =   44
            Left            =   2865
            TabIndex        =   30
            Top             =   2340
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
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   45
            Left            =   2880
            TabIndex        =   29
            Top             =   1425
            Width           =   2715
         End
         Begin VB.Line Line1 
            Index           =   26
            X1              =   4305
            X2              =   2865
            Y1              =   765
            Y2              =   765
         End
         Begin VB.Line Line1 
            Index           =   27
            X1              =   4305
            X2              =   2865
            Y1              =   1050
            Y2              =   1050
         End
         Begin VB.Line Line1 
            Index           =   28
            X1              =   4305
            X2              =   2865
            Y1              =   1350
            Y2              =   1350
         End
         Begin VB.Line Line1 
            Index           =   29
            X1              =   1605
            X2              =   165
            Y1              =   1965
            Y2              =   1965
         End
         Begin VB.Line Line1 
            Index           =   30
            X1              =   1620
            X2              =   180
            Y1              =   2265
            Y2              =   2265
         End
         Begin VB.Line Line1 
            Index           =   31
            X1              =   1635
            X2              =   195
            Y1              =   2565
            Y2              =   2565
         End
         Begin VB.Line Line1 
            Index           =   32
            X1              =   4290
            X2              =   2850
            Y1              =   1965
            Y2              =   1965
         End
         Begin VB.Line Line1 
            Index           =   33
            X1              =   4305
            X2              =   2865
            Y1              =   2265
            Y2              =   2265
         End
         Begin VB.Line Line1 
            Index           =   34
            X1              =   4305
            X2              =   2865
            Y1              =   2565
            Y2              =   2565
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah                Liter"
            Height          =   255
            Index           =   46
            Left            =   240
            TabIndex        =   28
            Top             =   2940
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Mulai"
            Height          =   255
            Index           =   47
            Left            =   225
            TabIndex        =   27
            Top             =   3240
            Width           =   2865
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Selesai"
            Height          =   255
            Index           =   48
            Left            =   210
            TabIndex        =   26
            Top             =   3525
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
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   49
            Left            =   225
            TabIndex        =   25
            Top             =   2625
            Width           =   2715
         End
         Begin VB.Line Line1 
            Index           =   38
            X1              =   1650
            X2              =   210
            Y1              =   3180
            Y2              =   3180
         End
         Begin VB.Line Line1 
            Index           =   39
            X1              =   1650
            X2              =   210
            Y1              =   3480
            Y2              =   3480
         End
         Begin VB.Line Line1 
            Index           =   40
            X1              =   1650
            X2              =   210
            Y1              =   3780
            Y2              =   3780
         End
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4125
         MaskColor       =   &H000000C0&
         Picture         =   "frmPbleacing.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "SPPH"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "desk_BLEACHING"
         DataSource      =   "DDE"
         Height          =   630
         Index           =   40
         Left            =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Tag             =   "bleaching"
         Top             =   5985
         Width           =   3540
      End
      Begin MSComCtl2.DTPicker tanggal 
         DataField       =   "tanggal_bleaching"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   0
         Left            =   2280
         TabIndex        =   75
         Tag             =   "bleaching"
         Top             =   405
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   58458115
         CurrentDate     =   39365
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   2385
         X2              =   345
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   2400
         X2              =   360
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2400
         X2              =   360
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2400
         X2              =   360
         Y1              =   405
         Y2              =   405
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tangki"
         Height          =   255
         Index           =   12
         Left            =   360
         TabIndex        =   80
         Top             =   1095
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   79
         Top             =   795
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
         Caption         =   "No Ekstrasi"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   77
         Top             =   165
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
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   50
         Left            =   300
         TabIndex        =   76
         Top             =   5730
         Width           =   2055
      End
   End
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   6855
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   1005
      BindFormTAG     =   "ext"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmPbleacing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents Bleacing  As frmCaller
Attribute Bleacing.VB_VarHelpID = -1
Dim rsbleacing As New DBQuick
Dim nom As Integer

Private Sub Bleacing_RowColChange(ByVal TagForm As String, _
                                  ByVal pRecordset As ADODB.Recordset)
    DDE.GetFieldByName("no_ekstrasi") = rsbleacing.DBRecordset.Fields("no_ekstrasi")
End Sub

Private Sub cmdLink_Click()
    rsbleacing.DBOpen "select * from BLEACHING", CNN

    If rsbleacing.DBRecordset.EOF Then
        rsbleacing.DBOpen "select * from ACID_TREATMEN", CNN
    Else
        rsbleacing.DBOpen "select * from BLEACHING, ACID_TREATMEN where BLEACHING.no_ekstrasi <> ACID_TREATMEN.no_ekstrasi ", CNN
    End If

    Set Bleacing = New frmCaller
    Set Bleacing.FormData = rsbleacing.DBRecordset
    Bleacing.FromTagActive = "BLEACHING TREATMEN"
    Bleacing.CaptionLink = "BLEACHING TREATMEN"
End Sub

Private Sub DDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb

        Case tmbAddNew:
            CmdLink.Enabled = True
            txt(0).Enabled = False
    End Select

End Sub

Private Sub DTPicker1_LostFocus()
    txt(nom).Text = Format(DTPicker1.value, "hh:mm:ss")
    DTPicker1.Visible = False
End Sub

Private Sub Form_Load()

    With DDE
        Set .BindForm = Me
        .BindFormTAG = "BLEACHING"
        Set .ActiveConnection = CNN
        .PrepareQuery = "select * from BLEACHING"
    End With

    'HiasForm Picture1, Me
    HiasFormManTell Picture2, Me
    seting Me
End Sub

Function Del()
    DDE.PrepareDelete = "delete  from BLEACHING where no_ekstrasi = '" + DDE.GetFieldByName("no_ekstrasi") + "'"
End Function

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

Function simpan()

    With DDE
        .PrepareAppend = "insert into BLEACHING (no_ekstrasi,tanggal_BLEACHING,Grup,tangki," & _
           " jumlah_air,jml_BLEACHING, konsentrasi_BLEACHING,jml_BLEACHING_akhir,konsentrasi_BLEACHING_akhir, suhu_BLEACHING_akhir," & _
           " waktu_mulai_treatmen,suhu,waktu_selesai_treatmen," & _
           " c_1_jumlah, c_1_mulai, c_1_selesai," & _
           " c_2_jumlah, c_2_mulai, c_2_selesai," & _
           " c_3_jumlah, c_3_mulai, c_3_selesai," & _
           " c_4_jumlah, c_4_mulai, c_4_selesai," & _
           " c_5_jumlah, c_5_mulai, c_5_selesai," & _
           " c_6_jumlah, c_6_mulai, c_6_selesai,desk_BLEACHING,ph) values" & _
           "('" + DDE.GetFieldByName("no_ekstrasi") + "','" + Format(tanggal(0).value, "yyyy-MM-dd") + "', " & _
           " '" + DDE.GetFieldByName("Grup") + "', '" + DDE.GetFieldByName("Tangki") + "', " & _
           " '" + DDE.GetFieldByName("jumlah_air") + "', " & _
           " '" + DDE.GetFieldByName("jml_BLEACHING") + "', '" + DDE.GetFieldByName("konsentrasi_BLEACHING") + "'," & _
           " '" + DDE.GetFieldByName("jml_BLEACHING_akhir") + "','" + DDE.GetFieldByName("konsentrasi_BLEACHING_akhir") + "', '" + DDE.GetFieldByName("suhu_BLEACHING_akhir") + "', " & _
           " '" + DDE.GetFieldByName("waktu_mulai_treatmen") + "', '" + DDE.GetFieldByName("suhu") + "', " & _
           " '" + DDE.GetFieldByName("waktu_selesai_treatmen") + "', " & _
           " '" + DDE.GetFieldByName("c_1_jumlah") + "','" + DDE.GetFieldByName("c_1_mulai") + "','" + DDE.GetFieldByName("c_1_selesai") + "', " & _
           " '" + DDE.GetFieldByName("c_2_jumlah") + "','" + DDE.GetFieldByName("c_2_mulai") + "','" + DDE.GetFieldByName("c_2_selesai") + "', " & _
           " '" + DDE.GetFieldByName("c_3_jumlah") + "','" + DDE.GetFieldByName("c_3_mulai") + "','" + DDE.GetFieldByName("c_3_selesai") + "', " & _
           " '" + DDE.GetFieldByName("c_4_jumlah") + "','" + DDE.GetFieldByName("c_4_mulai") + "','" + DDE.GetFieldByName("c_4_selesai") + "','" + DDE.GetFieldByName("c_5_jumlah") + "','" + DDE.GetFieldByName("c_5_mulai") + "','" + DDE.GetFieldByName("c_5_selesai") + "', " & _
           " '" + DDE.GetFieldByName("c_6_jumlah") + "','" + DDE.GetFieldByName("c_6_mulai") + "','" + DDE.GetFieldByName("c_6_selesai") + "','" + DDE.GetFieldByName("desk_BLEACHING") + "','" + DDE.GetFieldByName("ph") + "')"

        .PrepareUpdate = " update BLEACHING set tanggal_BLEACHING = '" & Format(tanggal(0).value, "yyyy-MM-dd") & "', Grup = '" & DDE.GetFieldByName("Grup") & "', tangki = '" & DDE.GetFieldByName("Tangki") & "',jumlah_air= '" & DDE.GetFieldByName("jumlah_air") & "', " & _
           " jml_BLEACHING= '" & DDE.GetFieldByName("jml_BLEACHING") & "', konsentrasi_BLEACHING= '" & DDE.GetFieldByName("konsentrasi_BLEACHING") & "', jml_BLEACHING_akhir= '" & DDE.GetFieldByName("jml_BLEACHING_akhir") & "', konsentrasi_BLEACHING_akhir = '" & DDE.GetFieldByName("konsentrasi_BLEACHING_akhir") & "', " & _
           " suhu_BLEACHING_akhir= '" & DDE.GetFieldByName("suhu_BLEACHING_akhir") & "', waktu_mulai_treatmen= '" & DDE.GetFieldByName("waktu_mulai_treatmen") & "', suhu= '" & DDE.GetFieldByName("suhu") & "', waktu_selesai_treatmen = '" & DDE.GetFieldByName("waktu_selesai_treatmen") & "', " & _
           " c_1_jumlah= '" & DDE.GetFieldByName("c_1_jumlah") & "', c_1_mulai= '" & DDE.GetFieldByName("c_1_mulai") & "', c_1_selesai= '" & DDE.GetFieldByName("c_1_selesai") & "', " & _
           " c_2_jumlah= '" & DDE.GetFieldByName("c_2_jumlah") & "', c_2_mulai= '" & DDE.GetFieldByName("c_2_mulai") & "', c_2_selesai= '" & DDE.GetFieldByName("c_2_selesai") & "', " & _
           " c_3_jumlah= '" & DDE.GetFieldByName("c_3_jumlah") & "', c_3_mulai= '" & DDE.GetFieldByName("c_3_mulai") & "', c_3_selesai= '" & DDE.GetFieldByName("c_3_selesai") & "', " & _
           " c_4_jumlah= '" & DDE.GetFieldByName("c_4_jumlah") & "', c_4_mulai= '" & DDE.GetFieldByName("c_4_mulai") & "', c_4_selesai= '" & DDE.GetFieldByName("c_4_selesai") & "', " & _
           " c_5_jumlah= '" & DDE.GetFieldByName("c_5_jumlah") & "', c_5_mulai= '" & DDE.GetFieldByName("c_5_mulai") & "', c_5_selesai= '" & DDE.GetFieldByName("c_5_selesai") & "', " & _
           " c_6_jumlah= '" & DDE.GetFieldByName("c_6_jumlah") & "', c_6_mulai= '" & DDE.GetFieldByName("c_6_mulai") & "', c_6_selesai= '" & DDE.GetFieldByName("c_6_selesai") & "', " & _
           " desk_BLEACHING= '" & DDE.GetFieldByName("desk_BLEACHING") & "', ph = '" & DDE.GetFieldByName("ph") & "' where no_ekstrasi ='" & .GetFieldByName("no_ekstrasi") & "'"

    End With

End Function

Private Sub txt_GotFocus(Index As Integer)
    txt(Index).BackColor = &H79BCFF

    Select Case Index

        Case 35, 34, 16, 15, 19, 18, 22, 21, 25, 24, 5, 4
            DTPicker1.Visible = True
            DTPicker1.Move txt(Index).Left, txt(Index).Top
            nom = txt(Index).Index
    End Select

End Sub

Private Sub txt_LostFocus(Index As Integer)
    txt(Index).BackColor = vbWhite
End Sub
