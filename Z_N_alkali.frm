VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{D7BB8F75-AC9E-4E80-A526-70EA20ACFD16}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form Z_N_alkali 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proses Alkali Treatment"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8865
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   8865
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      ForeColor       =   &H80000008&
      Height          =   5790
      Left            =   -15
      ScaleHeight     =   5760
      ScaleWidth      =   8865
      TabIndex        =   1
      Top             =   -15
      Width           =   8895
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "pre_lot_powder"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   1500
         TabIndex        =   49
         Tag             =   "powder"
         Top             =   120
         Width           =   1845
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "pre_lot_powder"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1500
         TabIndex        =   47
         Tag             =   "powder"
         Top             =   690
         Width           =   1845
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "pre_lot_powder"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1500
         TabIndex        =   46
         Tag             =   "powder"
         Top             =   960
         Width           =   1845
      End
      Begin VB.TextBox Text2 
         Height          =   300
         Left            =   8220
         TabIndex        =   2
         Top             =   5805
         Width           =   1725
      End
      Begin TabDlg.SSTab tab 
         Height          =   4260
         Left            =   810
         TabIndex        =   3
         Top             =   1425
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   7514
         _Version        =   393216
         Tabs            =   6
         Tab             =   5
         TabsPerRow      =   6
         TabHeight       =   520
         TabCaption(0)   =   "Page 1"
         TabPicture(0)   =   "Z_N_alkali.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame1(2)"
         Tab(0).Control(1)=   "Text1(2)"
         Tab(0).Control(2)=   "Text1(3)"
         Tab(0).Control(3)=   "Frame1(0)"
         Tab(0).Control(4)=   "Frame1(1)"
         Tab(0).Control(5)=   "Label1(3)"
         Tab(0).Control(6)=   "Line2(4)"
         Tab(0).Control(7)=   "Line2(5)"
         Tab(0).Control(8)=   "Label1(4)"
         Tab(0).ControlCount=   9
         TabCaption(1)   =   "Page 2"
         TabPicture(1)   =   "Z_N_alkali.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame1(3)"
         Tab(1).Control(1)=   "Frame1(4)"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Page 3"
         TabPicture(2)   =   "Z_N_alkali.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame1(5)"
         Tab(2).Control(1)=   "Frame1(6)"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Page 4"
         TabPicture(3)   =   "Z_N_alkali.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame1(7)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Page 5"
         TabPicture(4)   =   "Z_N_alkali.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame1(8)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Page 6"
         TabPicture(5)   =   "Z_N_alkali.frx":008C
         Tab(5).ControlEnabled=   -1  'True
         Tab(5).Control(0)=   "Frame1(9)"
         Tab(5).Control(0).Enabled=   0   'False
         Tab(5).ControlCount=   1
         Begin VB.Frame Frame1 
            Caption         =   "Kondisi"
            Height          =   690
            Index           =   1
            Left            =   -71400
            TabIndex        =   41
            Top             =   1350
            Width           =   2595
            Begin VB.OptionButton Option1 
               Caption         =   "Bersih"
               Height          =   240
               Index           =   3
               Left            =   195
               TabIndex        =   43
               Top             =   330
               Width           =   1095
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Kotor"
               Height          =   240
               Index           =   2
               Left            =   1320
               TabIndex        =   42
               Top             =   330
               Width           =   990
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Tempat Alkali Treatment"
            Height          =   690
            Index           =   0
            Left            =   -74610
            TabIndex        =   38
            Top             =   1350
            Width           =   2580
            Begin VB.OptionButton Option1 
               Caption         =   "Bak Luar"
               Height          =   240
               Index           =   1
               Left            =   1335
               TabIndex        =   40
               Top             =   300
               Width           =   990
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Reaktor"
               Height          =   240
               Index           =   0
               Left            =   195
               TabIndex        =   39
               Top             =   315
               Width           =   1095
            End
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            DataField       =   "pre_lot_powder"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   -72945
            TabIndex        =   37
            Tag             =   "powder"
            Top             =   945
            Width           =   1230
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            DataField       =   "pre_lot_powder"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   -72945
            TabIndex        =   36
            Tag             =   "powder"
            Top             =   675
            Width           =   1230
         End
         Begin VB.Frame Frame1 
            Caption         =   "Larutan Alkali Bekas"
            Height          =   1770
            Index           =   2
            Left            =   -74625
            TabIndex        =   34
            Top             =   2190
            Width           =   5835
            Begin MSFlexGridLib.MSFlexGrid AB 
               Height          =   1470
               Left            =   120
               TabIndex        =   35
               Top             =   240
               Width           =   5610
               _ExtentX        =   9895
               _ExtentY        =   2593
               _Version        =   393216
               Rows            =   1
               Cols            =   4
               FixedCols       =   0
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Air Bersih untuk penambahan proses alkali"
            Height          =   1770
            Index           =   3
            Left            =   -74790
            TabIndex        =   31
            Top             =   435
            Width           =   5835
            Begin VB.TextBox Text3 
               Height          =   375
               Left            =   3195
               TabIndex        =   33
               Top             =   1035
               Width           =   1875
            End
            Begin MSFlexGridLib.MSFlexGrid ab1 
               Height          =   1470
               Left            =   120
               TabIndex        =   32
               Top             =   240
               Width           =   5610
               _ExtentX        =   9895
               _ExtentY        =   2593
               _Version        =   393216
               Rows            =   1
               Cols            =   4
               FixedCols       =   0
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Alkali Baru"
            Height          =   1770
            Index           =   4
            Left            =   -74775
            TabIndex        =   28
            Top             =   2265
            Width           =   5835
            Begin VB.TextBox Text4 
               Height          =   375
               Left            =   3195
               TabIndex        =   30
               Top             =   1035
               Width           =   1875
            End
            Begin MSFlexGridLib.MSFlexGrid ab2 
               Height          =   1470
               Left            =   120
               TabIndex        =   29
               Top             =   240
               Width           =   5610
               _ExtentX        =   9895
               _ExtentY        =   2593
               _Version        =   393216
               Rows            =   1
               Cols            =   4
               FixedCols       =   0
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Larutan Alkali Akhir"
            Height          =   1770
            Index           =   5
            Left            =   -74880
            TabIndex        =   25
            Top             =   450
            Width           =   5835
            Begin VB.TextBox Text5 
               Height          =   375
               Left            =   3195
               TabIndex        =   27
               Top             =   1035
               Width           =   1875
            End
            Begin MSFlexGridLib.MSFlexGrid ab4 
               Height          =   1470
               Left            =   120
               TabIndex        =   26
               Top             =   240
               Width           =   5610
               _ExtentX        =   9895
               _ExtentY        =   2593
               _Version        =   393216
               Rows            =   1
               Cols            =   4
               FixedCols       =   0
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Memasukan Rumput Laut"
            Height          =   1770
            Index           =   6
            Left            =   -74850
            TabIndex        =   22
            Top             =   2250
            Width           =   5835
            Begin VB.TextBox Text6 
               Height          =   375
               Left            =   3195
               TabIndex        =   24
               Top             =   1035
               Width           =   1875
            End
            Begin MSFlexGridLib.MSFlexGrid ab5 
               Height          =   1470
               Left            =   120
               TabIndex        =   23
               Top             =   240
               Width           =   5610
               _ExtentX        =   9895
               _ExtentY        =   2593
               _Version        =   393216
               Rows            =   1
               Cols            =   4
               FixedCols       =   0
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Proses Alkali Treatment"
            Height          =   3435
            Index           =   7
            Left            =   -74820
            TabIndex        =   20
            Top             =   555
            Width           =   6060
            Begin MSFlexGridLib.MSFlexGrid ab6 
               Height          =   3045
               Left            =   120
               TabIndex        =   21
               Top             =   240
               Width           =   5790
               _ExtentX        =   10213
               _ExtentY        =   5371
               _Version        =   393216
               Rows            =   1
               Cols            =   4
               FixedCols       =   0
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Pemindahan rumput laut dari bak luar ke reaktor"
            Height          =   3570
            Index           =   8
            Left            =   -74790
            TabIndex        =   15
            Top             =   510
            Width           =   5835
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   330
               Index           =   0
               Left            =   1515
               TabIndex        =   16
               Top             =   615
               Width           =   1470
               _ExtentX        =   2593
               _ExtentY        =   582
               _Version        =   393216
               Format          =   58720258
               CurrentDate     =   39490
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   330
               Index           =   1
               Left            =   1515
               TabIndex        =   17
               Top             =   990
               Width           =   1470
               _ExtentX        =   2593
               _ExtentY        =   582
               _Version        =   393216
               Format          =   58720258
               CurrentDate     =   39490
            End
            Begin VB.Label Label2 
               Caption         =   "Waktu Mulai"
               Height          =   240
               Index           =   0
               Left            =   255
               TabIndex        =   19
               Top             =   675
               Width           =   1005
            End
            Begin VB.Label Label2 
               Caption         =   "Waktu Selesai"
               Height          =   240
               Index           =   1
               Left            =   255
               TabIndex        =   18
               Top             =   1020
               Width           =   1245
            End
            Begin VB.Line Line2 
               Index           =   6
               X1              =   195
               X2              =   2535
               Y1              =   930
               Y2              =   930
            End
            Begin VB.Line Line2 
               Index           =   7
               X1              =   180
               X2              =   2520
               Y1              =   1305
               Y2              =   1305
            End
         End
         Begin VB.Frame Frame1 
            Height          =   3660
            Index           =   9
            Left            =   120
            TabIndex        =   4
            Top             =   450
            Width           =   6180
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               DataField       =   "pre_lot_powder"
               DataSource      =   "DDE"
               Enabled         =   0   'False
               Height          =   285
               Index           =   4
               Left            =   2205
               TabIndex        =   7
               Tag             =   "powder"
               Top             =   750
               Width           =   615
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   2025
               TabIndex        =   6
               Text            =   "1"
               Top             =   330
               Width           =   825
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Add to Grid"
               Height          =   345
               Left            =   360
               TabIndex        =   5
               Top             =   1275
               Width           =   1080
            End
            Begin MSFlexGridLib.MSFlexGrid ab7 
               Height          =   1845
               Left            =   345
               TabIndex        =   8
               Top             =   1650
               Width           =   5505
               _ExtentX        =   9710
               _ExtentY        =   3254
               _Version        =   393216
               Rows            =   1
               Cols            =   3
               FixedCols       =   0
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   330
               Index           =   2
               Left            =   4530
               TabIndex        =   9
               Top             =   285
               Width           =   1470
               _ExtentX        =   2593
               _ExtentY        =   582
               _Version        =   393216
               CustomFormat    =   "hh:mm:ss"
               Format          =   58720258
               CurrentDate     =   39490
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   330
               Index           =   3
               Left            =   4530
               TabIndex        =   10
               Top             =   690
               Width           =   1470
               _ExtentX        =   2593
               _ExtentY        =   582
               _Version        =   393216
               Format          =   58720258
               CurrentDate     =   39490
            End
            Begin VB.Label Label2 
               Caption         =   "Waktu Mulai"
               Height          =   240
               Index           =   2
               Left            =   3540
               TabIndex        =   14
               Top             =   375
               Width           =   1005
            End
            Begin VB.Label Label2 
               Caption         =   "Waktu Selesai"
               Height          =   240
               Index           =   3
               Left            =   3390
               TabIndex        =   13
               Top             =   720
               Width           =   1245
            End
            Begin VB.Line Line2 
               Index           =   8
               X1              =   3375
               X2              =   5715
               Y1              =   630
               Y2              =   630
            End
            Begin VB.Line Line2 
               Index           =   9
               X1              =   3375
               X2              =   5715
               Y1              =   1005
               Y2              =   1005
            End
            Begin VB.Label Label2 
               Caption         =   "Jumlah air Pencucian(liter)"
               Height          =   240
               Index           =   4
               Left            =   285
               TabIndex        =   12
               Top             =   780
               Width           =   1920
            End
            Begin VB.Line Line2 
               Index           =   10
               X1              =   240
               X2              =   2580
               Y1              =   1020
               Y2              =   1020
            End
            Begin VB.Line Line2 
               Index           =   11
               X1              =   240
               X2              =   2850
               Y1              =   645
               Y2              =   645
            End
            Begin VB.Label Label3 
               Caption         =   "Proses Pencucian"
               Height          =   240
               Left            =   315
               TabIndex        =   11
               Top             =   375
               Width           =   1410
            End
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Berat rumput Laut(kg)"
            Height          =   255
            Index           =   4
            Left            =   -74625
            TabIndex        =   45
            Top             =   990
            Width           =   1710
         End
         Begin VB.Line Line2 
            Index           =   5
            X1              =   -74625
            X2              =   -72285
            Y1              =   1215
            Y2              =   1215
         End
         Begin VB.Line Line2 
            Index           =   4
            X1              =   -74625
            X2              =   -72285
            Y1              =   945
            Y2              =   945
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "No Stok Rumput Laut"
            Height          =   195
            Index           =   3
            Left            =   -74625
            TabIndex        =   44
            Top             =   720
            Width           =   1590
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   1500
         TabIndex        =   48
         Top             =   390
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd-MMMM-yyyy"
         Format          =   58720259
         CurrentDate     =   39489
         MaxDate         =   39489
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   420
         X2              =   2760
         Y1              =   390
         Y2              =   390
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Ekstraksi"
         Height          =   255
         Index           =   13
         Left            =   420
         TabIndex        =   53
         Top             =   135
         Width           =   1755
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   255
         Index           =   0
         Left            =   390
         TabIndex        =   52
         Top             =   465
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         Height          =   240
         Index           =   1
         Left            =   420
         TabIndex        =   51
         Top             =   735
         Width           =   750
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   420
         X2              =   2760
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanki"
         Height          =   255
         Index           =   2
         Left            =   420
         TabIndex        =   50
         Top             =   1020
         Width           =   570
      End
      Begin VB.Line Line2 
         Index           =   2
         X1              =   420
         X2              =   2760
         Y1              =   1230
         Y2              =   1230
      End
      Begin VB.Line Line2 
         Index           =   3
         X1              =   405
         X2              =   3315
         Y1              =   690
         Y2              =   690
      End
      Begin VB.Line Line1 
         X1              =   90
         X2              =   8010
         Y1              =   1350
         Y2              =   1350
      End
   End
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5805
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   1005
      BindFormTAG     =   "powder"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "Z_N_alkali"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim satuan As String
Dim DgDetail As MSFlexGrid

Sub Grid_header()
    On Error Resume Next
    Dim ctrl As Object

    For Each ctrl In Me.Controls

        If TypeOf ctrl Is MSFlexGrid Then

            With ctrl
                .Row = 0
                .Col = 0
                .ColWidth(0) = 2400
                .Col = 1
                .Text = "Reaktor"
                .ColWidth(1) = 1000
                .Col = 2
                .Text = "Bak Luar 1"
                .ColWidth(2) = 1000
                .Col = 3
                .Text = "Bak Luar 2"
                .ColWidth(3) = 1000
            End With

        End If

    Next ctrl

End Sub

Sub grid_alkali_akhir(grid As MSFlexGrid)

    With grid
        .AddItem "jumlah " & satuan
        .AddItem "Konsentrasi (%)"
        .AddItem "Suhu (c)"
    End With

End Sub

Sub rumput_laut(grid As MSFlexGrid)

    With grid
        .AddItem "Suhu (c)"
        .AddItem "Waktu"
    End With

End Sub

Sub grid_air(grid As MSFlexGrid)
    grid.AddItem "Jumlah " & satuan
End Sub

Sub grid_cuci()

    With ab7
        .Row = 0
        .Col = 0
        .Text = "Proses"
        .ColWidth(0) = 1500
        .Col = 1
        .Text = "Keterangan"
        .ColWidth(1) = 2000
        .Col = 2
        .Text = "Jumlah"
        .ColWidth(2) = 1500

    End With

End Sub

Sub Grid_Data(grid As MSFlexGrid)

    With grid
        .AddItem "Jumlah " & satuan
        .AddItem "Konsentrasi(%)"
        .AddItem "Waktu"
    End With

End Sub

Sub Grid_alkali_baru(grid As MSFlexGrid)
    grid.AddItem "Type Alkali"
    Grid_Data grid
End Sub

Private Sub AB_Click()
    Text2.SetFocus
    Set DgDetail = AB
    Text2.Text = ""
End Sub

Private Sub ab1_Click()
    Text2.SetFocus
    Set DgDetail = ab1
    Text2.Text = ""
End Sub

Private Sub ab2_Click()
    Text2.SetFocus
    Set DgDetail = ab2
    Text2.Text = ""
End Sub

Private Sub ab4_Click()
    Text2.SetFocus
    Set DgDetail = ab4
    Text2.Text = ""
End Sub

Private Sub ab5_Click()
    Text2.SetFocus
    Set DgDetail = ab5
    Text2.Text = ""
End Sub

Private Sub ab6_Click()
    Text2.SetFocus
    Set DgDetail = ab6
    Text2.Text = ""
End Sub

Private Sub Command1_Click()

    With ab7
        .MergeCol(0) = True
        .AddItem "Pencucian " & Combo1.Text + vbTab + "Jumlah Air Pencucian(liter)" + vbTab + Text1(4).Text
        .AddItem "Pencucian " & Combo1.Text + vbTab + "Waktu Mulai Pencucian" + vbTab + Text6.Text
        .AddItem "Pencucian " & Combo1.Text + vbTab + "Waktu Selesai Pencucian" + vbTab + Text5.Text
        .MergeCells = flexMergeRestrictColumns
    End With

    Combo1.Text = Val(Combo1.Text) + 1
End Sub

Private Sub DTPicker2_LostFocus(Index As Integer)

    Select Case Index

        Case 2
            Text6.Text = Format(DTPicker2(2).value, "hh:mm:ss")

        Case 3
            Text5.Text = Format(DTPicker2(2).value, "hh:mm:ss")
    End Select

End Sub

Private Sub Form_Load()
    Grid_header
    satuan = "(Liter)"
    Grid_Data AB
    grid_air ab1
    satuan = "(Kg)"
    Grid_alkali_baru ab2
    grid_alkali_akhir ab4
    rumput_laut ab5
    waktu ab6

    grid_cuci

    For I = 1 To 10
        Combo1.AddItem I
    Next I

End Sub

Sub waktu(grid As MSFlexGrid)

    With grid
        .AddItem "Waktu Mulai Treatment"
        .AddItem "Suhu Setelah 1 Jam"
        .AddItem "Suhu Setelah 2 Jam"
        .AddItem "Suhu Setelah 3 Jam"
        .AddItem "Suhu Setelah 4 Jam"
        .AddItem "Suhu Setelah 5 Jam"
        .AddItem "Suhu Setelah 6 Jam"
        .AddItem "Suhu Setelah 7 Jam"
        .AddItem "Suhu Setelah 8 Jam"
        .AddItem "Suhu Setelah 9 Jam"
        .AddItem "Suhu Setelah 10 Jam"
        .AddItem "Suhu Setelah 11 Jam"
        .AddItem "Suhu Setelah 12 Jam"
        .AddItem "Waktu Selesai Alkali Treatment"
    End With

End Sub

Private Sub Text2_Change()
    edit_grid
End Sub

Sub edit_grid()

    If DgDetail.Col = 0 Then
        Text2.Text = ""
        Exit Sub
    Else
        DgDetail.Text = Text2.Text
    End If

End Sub
