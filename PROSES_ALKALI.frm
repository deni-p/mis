VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{43E6F32B-2B03-46D3-9276-69426FE6D51B}#1.0#0"; "SemeruDC.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PROSES_ALKALI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ALKALI TREATMEN"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6075
      Index           =   0
      Left            =   75
      ScaleHeight     =   6045
      ScaleWidth      =   10275
      TabIndex        =   0
      Top             =   75
      Width           =   10305
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   5955
         Left            =   45
         ScaleHeight     =   5925
         ScaleWidth      =   10155
         TabIndex        =   1
         Top             =   45
         Width           =   10185
         Begin VB.Frame Frame2 
            BackColor       =   &H00EAAF6F&
            Caption         =   "Kondisi"
            Height          =   630
            Left            =   6675
            TabIndex        =   21
            Top             =   90
            Width           =   1950
            Begin VB.OptionButton Option1 
               Appearance      =   0  'Flat
               BackColor       =   &H00EAAF6F&
               Caption         =   "Bersih"
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   3
               Left            =   180
               TabIndex        =   23
               Top             =   285
               Width           =   765
            End
            Begin VB.OptionButton Option1 
               Appearance      =   0  'Flat
               BackColor       =   &H00EAAF6F&
               Caption         =   "Kotor"
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   2
               Left            =   1020
               TabIndex        =   22
               Top             =   285
               Width           =   780
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00EAAF6F&
            Caption         =   "Tempat Alkali Treatmen"
            Height          =   630
            Left            =   4275
            TabIndex        =   18
            Top             =   90
            Width           =   2310
            Begin VB.OptionButton Option1 
               Appearance      =   0  'Flat
               BackColor       =   &H00EAAF6F&
               Caption         =   "Bak Luar"
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   1
               Left            =   1170
               TabIndex        =   20
               Top             =   285
               Width           =   1005
            End
            Begin VB.OptionButton Option1 
               Appearance      =   0  'Flat
               BackColor       =   &H00EAAF6F&
               Caption         =   "Reaktor"
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   0
               Left            =   180
               TabIndex        =   19
               Top             =   285
               Width           =   915
            End
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "berat"
            DataSource      =   "DDE"
            Height          =   300
            Index           =   3
            Left            =   1965
            TabIndex        =   16
            Tag             =   "alkali"
            Top             =   1305
            Width           =   420
         End
         Begin TabDlg.SSTab ss 
            Height          =   3975
            Left            =   90
            TabIndex        =   12
            Top             =   1875
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   7011
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabHeight       =   520
            BackColor       =   15380335
            TabCaption(0)   =   "Page 1"
            TabPicture(0)   =   "PROSES_ALKALI.frx":0000
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "FLEX"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Page 2"
            TabPicture(1)   =   "PROSES_ALKALI.frx":001C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Label1(0)"
            Tab(1).Control(1)=   "Line1(5)"
            Tab(1).Control(2)=   "Line1(6)"
            Tab(1).Control(3)=   "Label1(1)"
            Tab(1).Control(4)=   "MFLEX"
            Tab(1).Control(5)=   "txt(4)"
            Tab(1).Control(6)=   "txt(5)"
            Tab(1).ControlCount=   7
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               DataField       =   "keterangan"
               DataSource      =   "DDE"
               Height          =   300
               Index           =   5
               Left            =   -70230
               TabIndex        =   26
               Tag             =   "alkali"
               Top             =   3525
               Width           =   5055
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               DataField       =   "ph_akhir"
               DataSource      =   "DDE"
               Height          =   300
               Index           =   4
               Left            =   -73050
               TabIndex        =   25
               Tag             =   "alkali"
               Top             =   3525
               Width           =   1695
            End
            Begin MSFlexGridLib.MSFlexGrid FLEX 
               Height          =   3570
               Left            =   30
               TabIndex        =   13
               Top             =   360
               Width           =   9885
               _ExtentX        =   17436
               _ExtentY        =   6297
               _Version        =   393216
               Rows            =   0
               Cols            =   4
               FixedRows       =   0
               FixedCols       =   0
               BackColorFixed  =   12632256
               GridColor       =   14737632
               AllowUserResizing=   3
               BorderStyle     =   0
            End
            Begin MSFlexGridLib.MSFlexGrid MFLEX 
               Height          =   3060
               Left            =   -74970
               TabIndex        =   14
               Top             =   360
               Width           =   6675
               _ExtentX        =   11774
               _ExtentY        =   5398
               _Version        =   393216
               Rows            =   1
               FixedRows       =   0
               FixedCols       =   0
               BackColorFixed  =   12632256
               BackColorBkg    =   -2147483633
               GridColor       =   14737632
               AllowUserResizing=   3
               BorderStyle     =   0
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Keterangan"
               Height          =   255
               Index           =   1
               Left            =   -71145
               TabIndex        =   27
               Top             =   3555
               Width           =   1740
            End
            Begin VB.Line Line1 
               Index           =   6
               X1              =   -69120
               X2              =   -71160
               Y1              =   3810
               Y2              =   3810
            End
            Begin VB.Line Line1 
               Index           =   5
               X1              =   -72885
               X2              =   -74925
               Y1              =   3810
               Y2              =   3810
            End
            Begin VB.Label Label1 
               Caption         =   "pH akhir alkali treatmen"
               Height          =   285
               Index           =   0
               Left            =   -74895
               TabIndex        =   24
               Top             =   3570
               Width           =   1830
            End
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "Tanki"
            DataSource      =   "DDE"
            Height          =   300
            Index           =   2
            Left            =   1965
            TabIndex        =   6
            Tag             =   "alkali"
            Top             =   1020
            Width           =   1695
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "Grup"
            DataSource      =   "DDE"
            Height          =   300
            Index           =   1
            Left            =   1965
            TabIndex        =   5
            Tag             =   "alkali"
            Top             =   735
            Width           =   1695
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "No_ekstraksi"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            ForeColor       =   &H00C0C0C0&
            Height          =   300
            Index           =   0
            Left            =   1965
            TabIndex        =   4
            Tag             =   "alkali"
            Top             =   150
            Width           =   2160
         End
         Begin VB.TextBox Text1 
            BorderStyle     =   0  'None
            Height          =   345
            Left            =   7650
            TabIndex        =   3
            Top             =   2790
            Width           =   1830
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "AT_tanggal"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   0
            Left            =   1950
            TabIndex        =   7
            Tag             =   "alkali"
            Top             =   420
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy"
            Format          =   16515075
            CurrentDate     =   39365
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   5250
            TabIndex        =   15
            Top             =   3105
            Width           =   1935
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Berat Rumput Laut (kg)"
            Height          =   255
            Index           =   4
            Left            =   255
            TabIndex        =   17
            Top             =   1365
            Width           =   1710
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   2220
            X2              =   180
            Y1              =   1590
            Y2              =   1590
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   2220
            X2              =   180
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   2220
            X2              =   180
            Y1              =   735
            Y2              =   735
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   2220
            X2              =   180
            Y1              =   435
            Y2              =   435
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   2235
            X2              =   195
            Y1              =   1305
            Y2              =   1305
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Tangki"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   11
            Top             =   1080
            Width           =   825
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Group"
            Height          =   255
            Index           =   2
            Left            =   255
            TabIndex        =   10
            Top             =   780
            Width           =   750
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   9
            Top             =   480
            Width           =   720
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "No Ekstrasi"
            Height          =   255
            Index           =   0
            Left            =   255
            TabIndex        =   8
            Top             =   210
            Width           =   975
         End
      End
   End
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   6240
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1005
      BindFormTAG     =   "powder"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "PROSES_ALKALI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim analisa As New DBQuick
Dim analisa_2 As New DBQuick
Dim kolom As Integer
Dim I As Integer
Dim brs As Integer

Private Function IndexAuto() As String
Dim Rc As New DBQuick
Dim TglSaiki As String
Dim Inom As Long
TglSaiki = Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2)
Rc.DBOpen "SELECT MAX(RIGHT(ID_ALKALI, 5)) AS MaxNom FROM [TBL_ALKALI_HEADER] WHERE (GETDATE() = { fn NOW() })", CNN, lckLockReadOnly
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
Dim tempat As String
Dim kondisi As String
Dim ID As String
ID = IndexAuto
If Option1(0).value = True Then tempat = "Reaktor"
If Option1(1).value = True Then tempat = "Bak Luar"
If Option1(3).value = True Then kondisi = "Bersih"
If Option1(2).value = True Then kondisi = "Kotor"
With DDE
.PrepareAppend = "insert into TBL_ALKALI_HEADER (id_alkali,no_ekstraksi,tanggal_alkali,grup,tanki,berat,tempat,kondisi,ph_akhir,keterangan) " & _
                 "values('" & ID & "', '" & txt(0).Text & "', '" & Format(DTPicker1(0).value, "yyyy-MM-dd") & "','" & txt(1).Text & "', '" & txt(2).Text & "', '" & txt(3).Text & "', '" & tempat & "', '" & kondisi & "', '" & txt(4).Text & "', '" & txt(5).Text & "')"

End With
End Function

Private Sub DDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If DDE.GetFieldByName("tempat") = "Reaktor" Then Option1(0).value = True
If DDE.GetFieldByName("tempat") = "Bak Luar" Then Option1(1).value = True
If DDE.GetFieldByName("kondisi") = "Bersih" Then Option1(3).value = True
If DDE.GetFieldByName("kondisi") = "Kotor" Then Option1(2).value = True

End Sub

Private Sub DDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
Case tmbSave
    DDE.IsChildMemberReady = True
    simpan
    
End Select

End Sub

Function simpan_detail()
Dim baris_data As Integer
Dim kolom_data As Integer

For baris_data = 0 To FLEX.Rows - 1
FLEX.Row = baris_data
     For kolom_data = 0 To FLEX.Cols - 1
     FLEX.Col = kolom_data
     DDE.PrepareAppend = "insert into TBL_ALKALI_DETAIL_PROSES (ID_alkali,proses,reaktor,bak_luar_1,bak_luar_2) values ('" & FLEX.Text & "' "
     Next kolom_data
Next baris_data

End Function


Private Sub FLEX_Click()
Text1.Text = ""
Text1.SetFocus
End Sub

Private Sub Form_Load()

With FLEX
.Row = 0
.ColWidth(0) = 6000
.ColWidth(1) = 1200
.ColWidth(2) = 1200
.ColWidth(3) = 1200
End With

With MFLEX
.Row = 0
.ColWidth(0) = 5000
.ColWidth(1) = 1400
End With


analisa.DBOpen "select * from  VIEW_PROSES WHERE id_form = 'FR01' order by id asc", CNN
analisa.DBRecordset.MoveFirst
I = 1
baris = 1
Do While Not analisa.DBRecordset.EOF
  If analisa.DBRecordset("status") = "Header" Then
     FLEX.AddItem analisa.DBRecordset("nama_analisa") + vbTab + "Reaktor" + vbTab + "Bak Luar 1" + vbTab + "Bak Luar 2"
     grid_warna FLEX, I, 3, True
  Else
     FLEX.AddItem analisa.DBRecordset("nama_analisa")
     grid_warna FLEX, I, 3, False
  End If
I = I + 1
baris = baris + 1
analisa.DBRecordset.MoveNext
Loop

analisa_2.DBOpen "select * from  VIEW_PROSES WHERE id_form = 'FR02' order by id asc", CNN
analisa_2.DBRecordset.MoveFirst
I = 1
Do While Not analisa_2.DBRecordset.EOF
  If analisa_2.DBRecordset("status") = "Header" Then
     MFLEX.AddItem analisa_2.DBRecordset("nama_analisa") + vbTab + "Waktu/Jumlah"
     grid_warna MFLEX, I, 1, True
  Else
     MFLEX.AddItem analisa_2.DBRecordset("nama_analisa")
     grid_warna MFLEX, I, 1, False
  End If
I = I + 1
analisa_2.DBRecordset.MoveNext
Loop

With DDE
Set .BindForm = Me
    .BindFormTAG = "alkali"
Set .ActiveConnection = CNN
    .PrepareQuery = " select * from tbl_alkali_header  "
End With
HiasForm Picture2, Me
seting Me


End Sub

Private Sub MFLEX_Click()
Text2.SetFocus
End Sub

Private Sub Text1_Change()
FLEX.Text = Text1.Text
End Sub

Private Sub Text1_DblClick()


FLEX.Row = baris
FLEX.Col = 1
m.Text = asa
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text1.Text = ""
End If
End Sub

Private Sub Text2_Change()
MFLEX.Text = Text2.Text
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.Text = ""
End If
End Sub
