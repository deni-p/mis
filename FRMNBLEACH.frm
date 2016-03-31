VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{D7BB8F75-AC9E-4E80-A526-70EA20ACFD16}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FRMNBLEACH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BLEACHING TREATMENT"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8775
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   8775
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6090
      Index           =   0
      Left            =   150
      ScaleHeight     =   6060
      ScaleWidth      =   8475
      TabIndex        =   0
      Top             =   195
      Width           =   8505
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   5925
         Left            =   105
         ScaleHeight     =   5895
         ScaleWidth      =   8160
         TabIndex        =   1
         Top             =   45
         Width           =   8190
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "no_ekstraksi"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   16
            Left            =   1500
            TabIndex        =   50
            Tag             =   "bleaching"
            Top             =   105
            Width           =   1845
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "grup"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   1500
            TabIndex        =   49
            Tag             =   "bleaching"
            Top             =   705
            Width           =   2460
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "tanki"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            Height          =   270
            Index           =   1
            Left            =   1500
            TabIndex        =   48
            Tag             =   "bleaching"
            Top             =   1005
            Width           =   2460
         End
         Begin VB.TextBox Text2 
            Height          =   300
            Left            =   8220
            TabIndex        =   47
            Top             =   5805
            Width           =   1725
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "Tanggal"
            DataSource      =   "DDE"
            Height          =   330
            Left            =   1485
            TabIndex        =   2
            Top             =   390
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   582
            _Version        =   393216
            Format          =   58982401
            CurrentDate     =   39493
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   4230
            Left            =   765
            TabIndex        =   3
            Top             =   1560
            Width           =   6645
            _ExtentX        =   11721
            _ExtentY        =   7461
            _Version        =   393216
            Tab             =   2
            TabHeight       =   520
            TabCaption(0)   =   "Page 1"
            TabPicture(0)   =   "FRMNBLEACH.frx":0000
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Text1(2)"
            Tab(0).Control(1)=   "Frame2(0)"
            Tab(0).Control(2)=   "Frame2(1)"
            Tab(0).Control(3)=   "Text1(15)"
            Tab(0).Control(4)=   "Label4"
            Tab(0).Control(5)=   "Label1(3)"
            Tab(0).Control(6)=   "Label1(16)"
            Tab(0).ControlCount=   7
            TabCaption(1)   =   "Page 2"
            TabPicture(1)   =   "FRMNBLEACH.frx":001C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame2(2)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Page 3"
            TabPicture(2)   =   "FRMNBLEACH.frx":0038
            Tab(2).ControlEnabled=   -1  'True
            Tab(2).Control(0)=   "Frame1(9)"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).ControlCount=   1
            Begin VB.Frame Frame1 
               Height          =   3660
               Index           =   9
               Left            =   240
               TabIndex        =   31
               Top             =   390
               Width           =   6225
               Begin VB.TextBox Text1 
                  Appearance      =   0  'Flat
                  DataField       =   "pre_lot_powder"
                  DataSource      =   "DDE"
                  Height          =   285
                  Index           =   4
                  Left            =   2205
                  TabIndex        =   34
                  Tag             =   "acid"
                  Top             =   750
                  Width           =   615
               End
               Begin VB.ComboBox Combo1 
                  Height          =   315
                  Left            =   2025
                  TabIndex        =   33
                  Tag             =   "acid"
                  Text            =   "1"
                  Top             =   330
                  Width           =   825
               End
               Begin VB.CommandButton Command1 
                  Caption         =   "Add to Grid"
                  Height          =   345
                  Left            =   360
                  TabIndex        =   32
                  Top             =   1275
                  Width           =   1080
               End
               Begin MSFlexGridLib.MSFlexGrid ab7 
                  Height          =   1845
                  Left            =   345
                  TabIndex        =   35
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
                  TabIndex        =   36
                  Tag             =   "acid"
                  Top             =   315
                  Width           =   1470
                  _ExtentX        =   2593
                  _ExtentY        =   582
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   58982402
                  CurrentDate     =   39490
               End
               Begin MSComCtl2.DTPicker DTPicker2 
                  Height          =   330
                  Index           =   3
                  Left            =   4530
                  TabIndex        =   37
                  Tag             =   "acid"
                  Top             =   690
                  Width           =   1470
                  _ExtentX        =   2593
                  _ExtentY        =   582
                  _Version        =   393216
                  Format          =   58982402
                  CurrentDate     =   39490
               End
               Begin VB.TextBox Text3 
                  Height          =   315
                  Index           =   0
                  Left            =   1980
                  TabIndex        =   38
                  Top             =   1785
                  Width           =   1365
               End
               Begin VB.TextBox Text3 
                  Height          =   315
                  Index           =   1
                  Left            =   3405
                  TabIndex        =   39
                  Top             =   1785
                  Width           =   1365
               End
               Begin VB.Label Label2 
                  Caption         =   "Waktu Mulai"
                  Height          =   240
                  Index           =   2
                  Left            =   3540
                  TabIndex        =   43
                  Top             =   375
                  Width           =   1005
               End
               Begin VB.Label Label2 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Waktu Selesai"
                  Height          =   240
                  Index           =   3
                  Left            =   3390
                  TabIndex        =   42
                  Top             =   780
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
                  TabIndex        =   41
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
                  TabIndex        =   40
                  Top             =   375
                  Width           =   1410
               End
            End
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               DataField       =   "Jumlah_air"
               DataSource      =   "DDE"
               Enabled         =   0   'False
               Height          =   285
               Index           =   2
               Left            =   -73710
               TabIndex        =   30
               Tag             =   "bleaching"
               Top             =   855
               Width           =   975
            End
            Begin VB.Frame Frame2 
               Caption         =   "Bleaching Agent"
               Height          =   1590
               Index           =   0
               Left            =   -74790
               TabIndex        =   21
               Top             =   1320
               Width           =   3030
               Begin VB.TextBox Text1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  DataField       =   "pre_lot_powder"
                  DataSource      =   "DDE"
                  Enabled         =   0   'False
                  Height          =   285
                  Index           =   3
                  Left            =   1380
                  TabIndex        =   25
                  Tag             =   "bleaching"
                  Top             =   255
                  Width           =   1470
               End
               Begin VB.TextBox Text1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  DataField       =   "pre_lot_powder"
                  DataSource      =   "DDE"
                  Enabled         =   0   'False
                  Height          =   285
                  Index           =   5
                  Left            =   1380
                  TabIndex        =   24
                  Tag             =   "bleaching"
                  Top             =   570
                  Width           =   1470
               End
               Begin VB.TextBox Text1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  DataField       =   "pre_lot_powder"
                  DataSource      =   "DDE"
                  Enabled         =   0   'False
                  Height          =   285
                  Index           =   6
                  Left            =   1380
                  TabIndex        =   23
                  Tag             =   "bleaching"
                  Top             =   885
                  Width           =   1470
               End
               Begin VB.TextBox Text1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  DataField       =   "pre_lot_powder"
                  DataSource      =   "DDE"
                  Enabled         =   0   'False
                  Height          =   285
                  Index           =   7
                  Left            =   1380
                  TabIndex        =   22
                  Tag             =   "bleaching"
                  Top             =   1200
                  Width           =   1470
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Type"
                  Height          =   240
                  Index           =   4
                  Left            =   570
                  TabIndex        =   29
                  Top             =   315
                  Width           =   810
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Jumlah(kg)"
                  Height          =   240
                  Index           =   5
                  Left            =   540
                  TabIndex        =   28
                  Top             =   615
                  Width           =   810
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Konsentrasi (%)"
                  Height          =   240
                  Index           =   6
                  Left            =   210
                  TabIndex        =   27
                  Top             =   930
                  Width           =   1155
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Waktu"
                  Height          =   240
                  Index           =   7
                  Left            =   810
                  TabIndex        =   26
                  Top             =   1245
                  Width           =   525
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "Larutan Bleaching Akhir"
               Height          =   1590
               Index           =   1
               Left            =   -71580
               TabIndex        =   14
               Top             =   1320
               Width           =   3030
               Begin VB.TextBox Text1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  DataField       =   "pre_lot_powder"
                  DataSource      =   "DDE"
                  Enabled         =   0   'False
                  Height          =   285
                  Index           =   9
                  Left            =   1380
                  TabIndex        =   17
                  Tag             =   "bleaching"
                  Top             =   885
                  Width           =   1470
               End
               Begin VB.TextBox Text1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  DataField       =   "pre_lot_powder"
                  DataSource      =   "DDE"
                  Enabled         =   0   'False
                  Height          =   285
                  Index           =   10
                  Left            =   1380
                  TabIndex        =   16
                  Tag             =   "bleaching"
                  Top             =   570
                  Width           =   1470
               End
               Begin VB.TextBox Text1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  DataField       =   "pre_lot_powder"
                  DataSource      =   "DDE"
                  Enabled         =   0   'False
                  Height          =   285
                  Index           =   11
                  Left            =   1380
                  TabIndex        =   15
                  Tag             =   "bleaching"
                  Top             =   255
                  Width           =   1470
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Suhu"
                  Height          =   240
                  Index           =   8
                  Left            =   915
                  TabIndex        =   20
                  Top             =   915
                  Width           =   405
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Konsentrasi (%)"
                  Height          =   240
                  Index           =   9
                  Left            =   225
                  TabIndex        =   19
                  Top             =   615
                  Width           =   1155
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Jumlah(kg)"
                  Height          =   240
                  Index           =   10
                  Left            =   540
                  TabIndex        =   18
                  Top             =   300
                  Width           =   810
               End
            End
            Begin VB.Frame Frame2 
               Height          =   1755
               Index           =   2
               Left            =   -74550
               TabIndex        =   5
               Top             =   1245
               Width           =   5655
               Begin VB.TextBox Text1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  DataField       =   "pre_lot_powder"
                  DataSource      =   "DDE"
                  Enabled         =   0   'False
                  Height          =   285
                  Index           =   14
                  Left            =   3075
                  TabIndex        =   9
                  Tag             =   "bleaching"
                  Top             =   885
                  Width           =   1470
               End
               Begin VB.TextBox Text1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  DataField       =   "pre_lot_powder"
                  DataSource      =   "DDE"
                  Enabled         =   0   'False
                  Height          =   285
                  Index           =   13
                  Left            =   3075
                  TabIndex        =   8
                  Tag             =   "bleaching"
                  Top             =   1215
                  Width           =   1470
               End
               Begin VB.TextBox Text1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  DataField       =   "pre_lot_powder"
                  DataSource      =   "DDE"
                  Enabled         =   0   'False
                  Height          =   285
                  Index           =   12
                  Left            =   3075
                  TabIndex        =   7
                  Tag             =   "bleaching"
                  Top             =   555
                  Width           =   1470
               End
               Begin VB.TextBox Text1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  DataField       =   "pre_lot_powder"
                  DataSource      =   "DDE"
                  Enabled         =   0   'False
                  Height          =   285
                  Index           =   8
                  Left            =   3075
                  TabIndex        =   6
                  Tag             =   "bleaching"
                  Top             =   225
                  Width           =   1470
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total Waktu Bleaching Treatment"
                  Height          =   240
                  Index           =   14
                  Left            =   450
                  TabIndex        =   13
                  Top             =   1260
                  Width           =   2550
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Suhu Pada 20 Menit"
                  Height          =   240
                  Index           =   12
                  Left            =   1410
                  TabIndex        =   12
                  Top             =   600
                  Width           =   1515
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Waktu Mulai Bleaching Treatment"
                  Height          =   240
                  Index           =   11
                  Left            =   465
                  TabIndex        =   11
                  Top             =   270
                  Width           =   2655
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Waktu Selesai Bleaching Treatment"
                  Height          =   240
                  Index           =   15
                  Left            =   315
                  TabIndex        =   10
                  Top             =   945
                  Width           =   2595
               End
            End
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               DataField       =   "keterangan"
               DataSource      =   "DDE"
               Enabled         =   0   'False
               Height          =   705
               Index           =   15
               Left            =   -73650
               MultiLine       =   -1  'True
               TabIndex        =   4
               Tag             =   "bleaching"
               Top             =   3150
               Width           =   2610
            End
            Begin VB.Label Label4 
               Caption         =   "Air Bersih untuk penambahan Proses Bleaching"
               Height          =   255
               Left            =   -74805
               TabIndex        =   46
               Top             =   480
               Width           =   3810
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Jumlah (liter)"
               Height          =   240
               Index           =   3
               Left            =   -74655
               TabIndex        =   45
               Top             =   900
               Width           =   945
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Keterangan"
               Height          =   240
               Index           =   16
               Left            =   -74730
               TabIndex        =   44
               Top             =   3195
               Width           =   945
            End
         End
         Begin VB.Line Line2 
            Index           =   1
            X1              =   420
            X2              =   3330
            Y1              =   390
            Y2              =   390
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Ekstraksi"
            Height          =   255
            Index           =   13
            Left            =   435
            TabIndex        =   54
            Top             =   135
            Width           =   1755
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal"
            Height          =   255
            Index           =   0
            Left            =   420
            TabIndex        =   53
            Top             =   450
            Width           =   1065
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Group"
            Height          =   240
            Index           =   1
            Left            =   435
            TabIndex        =   52
            Top             =   750
            Width           =   570
         End
         Begin VB.Line Line2 
            Index           =   0
            X1              =   420
            X2              =   3945
            Y1              =   990
            Y2              =   990
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanki"
            Height          =   255
            Index           =   2
            Left            =   420
            TabIndex        =   51
            Top             =   1020
            Width           =   570
         End
         Begin VB.Line Line2 
            Index           =   2
            X1              =   420
            X2              =   3945
            Y1              =   1275
            Y2              =   1275
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
            Y1              =   1410
            Y2              =   1410
         End
      End
   End
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   55
      Top             =   6525
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   1005
      BindFormTAG     =   "powder"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FRMNBLEACH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsbleaching As New ADODB.Recordset
Dim rsProses As New ADODB.Recordset
Dim rsCuci As New ADODB.Recordset
Dim browse As Boolean

Private Sub tab_DblClick()

End Sub

Private Sub Command1_Click()

    With ab7
        .MergeCol(0) = True
        .AddItem "Pencucian " & Combo1.Text + vbTab + "Jumlah Air Pencucian(liter)" + vbTab + Text1(4).Text
        .AddItem "Pencucian " & Combo1.Text + vbTab + "Waktu Mulai Pencucian" + vbTab + Text3(0).Text
        .AddItem "Pencucian " & Combo1.Text + vbTab + "Waktu Selesai Pencucian" + vbTab + Text3(1).Text
        .MergeCells = flexMergeRestrictColumns
    End With

    Combo1.Text = Val(Combo1.Text) + 1
End Sub

Private Sub DDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb

        Case tmbAddNew:
            Text1(16).SetFocus
            ab7.Rows = 1
            bersih
    End Select
      
End Sub

Private Sub DDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
                             ByVal pError As ADODB.Error, _
                             adStatus As ADODB.EventStatusEnum, _
                             ByVal pRecordset As ADODB.Recordset)
    Tampil_data
End Sub

Private Sub DDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb

        Case tmbSave:
            DDE.IsChildMemberReady = True
            simpan
            Simpan_proses_1
            simpan_proses_cuci
            browse = True

        Case tmbEdit:

        Case tmbDelete:
    End Select

End Sub

Function Tampil_data()
    Set rsbleaching = CNN.Execute("select * from Z_bleaching_Treatment where no_ekstraksi = '" & DDE.GetFieldByName("no_ekstraksi") & "'")

    If rsbleaching.EOF Then Exit Function
    Text1(2).Text = rsbleaching.Fields(1)
    Text1(3).Text = rsbleaching.Fields(2)
    Text1(5).Text = rsbleaching.Fields(3)
    Text1(6).Text = rsbleaching.Fields(4)
    Text1(7).Text = rsbleaching.Fields(5)
    Text1(11).Text = rsbleaching.Fields(6)
    Text1(10).Text = rsbleaching.Fields(7)
    Text1(9).Text = rsbleaching.Fields(8)
    Text1(8).Text = rsbleaching.Fields(9)
    Text1(12).Text = rsbleaching.Fields(10)
    Text1(14).Text = rsbleaching.Fields(11)
    Text1(13).Text = rsbleaching.Fields(12)
    Set rsCuci = CNN.Execute("select * from z_proses_cuci where no_ekstraksi = '" & DDE.GetFieldByName("no_ekstraksi") & "' and ket = '" & "Bleaching" & "'")

    If rsCuci.EOF Then Exit Function

    With ab7
        .Clear
        .Rows = 1
        rsCuci.MoveFirst

        Do While Not rsCuci.EOF
            .AddItem rsCuci.Fields(1) + vbTab + rsCuci.Fields(2) + vbTab + rsCuci.Fields(3)
            rsCuci.MoveNext
        Loop

    End With

End Function

Private Sub DTPicker2_LostFocus(Index As Integer)

    Select Case Index

        Case 2
            Text3(0).Text = Format(DTPicker2(2).value, "hh:mm:ss")

        Case 3
            Text3(1).Text = Format(DTPicker2(3).value, "hh:mm:ss")
    End Select

End Sub

Function simpan()
    Set rsbleaching = CNN.Execute("insert into Z_proses (no_ekstraksi,tanggal,grup,tanki,keterangan,proses) values ('" & Text1(16).Text & "', '" & DTPicker1.value & "', '" & Text1(0).Text & "', '" & Text1(1).Text & "', '" & Text1(15).Text & "','" & "bleachingTreatment" & "')")
End Function

Function update_data()
    Set rsbleaching = CNN.Execute("update  Z_proses set tanggal = '" & DTPicker1.value & "', grup = '" & Text1(0).Text & "', tanki = '" & Text1(1).Text & "', keterangan = '" & Text1(15).Text & "' where no_ekstraksi = '" & DDE.GetFieldByName("no_ekstraksi") & "'")
End Function

Function Simpan_proses_1()

    Set rsProses = CNN.Execute("insert into Z_bleaching_treatment(no_ekstraksi,jumlah_air,type_bleaching,jumlah_bleaching,konsentrasi_bleaching,waktu,jumlah_bleaching_akhir,konsentrasi_bleaching_akhir,suhu_1,waktu_mulai,suhu_2,waktu_selesai,total) values " & "('" & Text1(16).Text & "', '" & Text1(2).Text & "', '" & Text1(3).Text & "', '" & Text1(5).Text & "', '" & Text1(6).Text & "', '" & Text1(7).Text & "', '" & Text1(11).Text & "','" & Text1(10).Text & "', '" & Text1(9).Text & "', '" & Text1(8).Text & "', '" & Text1(12).Text & "', '" & Text1(14).Text & "', '" & Text1(13).Text & "')")

End Function

Function simpan_proses_cuci()
    Dim baris As Integer

    For baris = 1 To ab7.Rows - 1
        Set rsCuci = CNN.Execute("insert into Z_PROSES_CUCI(no_ekstraksi,keterangan, proses, nilai,ket) values ('" & Text1(16).Text & "','" & ab7.TextMatrix(baris, 0) & "','" & ab7.TextMatrix(baris, 1) & "','" & ab7.TextMatrix(baris, 2) & "','" & "Bleaching" & "')")
    Next baris

End Function

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
        .MergeCol(0) = True
        .MergeCells = flexMergeRestrictColumns
    End With

End Sub

Private Sub Form_Load()

    With DDE
        Set .BindForm = Me
        .BindFormTAG = "bleaching"
        Set .ActiveConnection = CNN
        .PrepareQuery = " select * from Z_proses where proses = '" & "bleachingTreatment" & "'"
    End With

    grid_cuci

    For I = 1 To 10
        Combo1.AddItem I
    Next I

    browse = True
End Sub

Function bersih()
    On Error Resume Next
    Dim ctrl As Control

    For Each ctrl In Me.Controls

        If TypeOf ctrl Is TextBox Then
            ctrl.Text = ""
        End If

    Next ctrl

End Function

