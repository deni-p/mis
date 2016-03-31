VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{43E6F32B-2B03-46D3-9276-69426FE6D51B}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form Bahan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pemakaian Bahan Di Produksi"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8985
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3660
      Left            =   135
      ScaleHeight     =   3630
      ScaleWidth      =   8730
      TabIndex        =   0
      Top             =   90
      Width           =   8760
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   3525
         Left            =   60
         ScaleHeight     =   3495
         ScaleWidth      =   8595
         TabIndex        =   1
         Top             =   60
         Width           =   8625
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "id_bahan"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   0
            Left            =   7110
            TabIndex        =   8
            Tag             =   "bahan"
            Top             =   525
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "grup"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   1
            Left            =   2070
            TabIndex        =   3
            Tag             =   "bahan"
            Top             =   735
            Width           =   1695
         End
         Begin MSDataGridLib.DataGrid DGDETAIL 
            Bindings        =   "Bahan.frx":0000
            Height          =   2175
            Left            =   120
            TabIndex        =   2
            Tag             =   "bahan"
            Top             =   1200
            Width           =   8310
            _ExtentX        =   14658
            _ExtentY        =   3836
            _Version        =   393216
            AllowUpdate     =   -1  'True
            AllowArrows     =   -1  'True
            Appearance      =   0
            DefColWidth     =   6667
            HeadLines       =   1
            RowHeight       =   15
            WrapCellPointer =   -1  'True
            RowDividerStyle =   1
            FormatLocked    =   -1  'True
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "DAFTAR PRODUK"
            ColumnCount     =   5
            BeginProperty Column00 
               DataField       =   "bagian"
               Caption         =   "Bagian"
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
               DataField       =   "nama_barang"
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
               DataField       =   "jumlah"
               Caption         =   "Jumlah"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "satuan"
               Caption         =   "satuan"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "keterangan"
               Caption         =   "keterangan"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  Alignment       =   3
                  DividerStyle    =   4
                  Button          =   -1  'True
                  WrapText        =   -1  'True
                  ColumnWidth     =   1514,835
               EndProperty
               BeginProperty Column01 
               EndProperty
               BeginProperty Column02 
               EndProperty
               BeginProperty Column03 
               EndProperty
               BeginProperty Column04 
               EndProperty
            EndProperty
         End
         Begin MSComCtl2.DTPicker tgl 
            DataField       =   "tanggal_ekstrasi"
            DataSource      =   "DDE"
            Height          =   315
            Left            =   2070
            TabIndex        =   4
            Tag             =   "bahan"
            Top             =   420
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy"
            Format          =   16384003
            CurrentDate     =   39365
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   495
            Width           =   2055
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Group"
            Height          =   255
            Index           =   2
            Left            =   135
            TabIndex        =   5
            Top             =   795
            Width           =   2055
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   2160
            X2              =   120
            Y1              =   735
            Y2              =   735
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   2160
            X2              =   120
            Y1              =   1035
            Y2              =   1035
         End
      End
   End
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   7
      Tag             =   "bahan"
      Top             =   3840
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   1005
      BindFormTAG     =   "cruz"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "Bahan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim id_bhn As String

Private Sub DDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
Case tmbAddNew
   id_bhn = IndexAuto
End Select
End Sub

Private Sub DDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim rsdetail As DBQuick
Set rsdetail = New DBQuick
rsdetail.DBOpen "select * from view_bahan where id_bahan = '" & DDE.GetFieldByName("id_bahan") & "'", CNN
Set DDE.ChildRecordset = rsdetail.DBRecordset.Clone(adLockBatchOptimistic)
Set DGdetail.DataSource = DDE.ChildRecordset
End Sub

Private Sub DDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
Case tmbSave
    DDE.IsChildMemberReady = True
    simpan_header
    simpan_detail
End Select
End Sub

Private Sub DgDetail_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   DDE.ChildRecordset.AddNew
End If
End Sub

Private Sub Form_Load()
With DDE
Set .BindForm = Me
    .BindFormTAG = "bahan"
Set .ActiveConnection = CNN
    .PrepareQuery = "select * from bahan_header"
End With
HiasForm Picture1, Me
seting Me
End Sub

Function simpan_header()

txt(0).Text = IndexAuto
DDE.PrepareAppend = "insert into bahan_header (id_bahan,grup,tanggal) values ('" & id_bhn & "','" & txt(1).Text & "', '" & Format(tgl.value, "yyyy-MM-dd") & "')"
DDE.PrepareUpdate = "update bahan_header set grup = '" & txt(1).Text & "', tanggal = '" & Format(tgl.value, "yyyy-MM-dd") & "' where id_bahan = '" & txt(0).Text & "'"

End Function
Function simpan_detail()
With DDE.ChildRecordset
   If .Recordcount <> 0 Then
       .MoveFirst
         If SendDataToServer(" delete from [bahan_detail] where (id_bahan = '" & DDE.GetFieldByName("id_bahan") & "')") = True Then
         Do
           If .EOF = True Then Exit Do
           SendDataToServer "insert into bahan_detail (id_bahan,bagian,nama_barang,jumlah,satuan,keterangan)  " & _
           " values ('" & txt(0).Text & "', " & _
           " '" & .Fields("bagian") & "', " & _
           " '" & .Fields("nama_barang") & "', " & _
           " '" & .Fields("jumlah") & "', " & _
           " '" & .Fields("satuan") & "', " & _
           " '" & .Fields("keterangan") & "')"
          .MoveNext
        Loop
        End If
        .MoveLast
        DGdetail.Refresh
        End If
    End With
End Function



Private Function IndexAuto() As String
Dim Rc As New DBQuick
Dim TglSaiki As String
Dim Inom As Long
TglSaiki = Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2)
Rc.DBOpen "SELECT MAX(RIGHT(id_bahan, 5)) AS MaxNom FROM [bahan_header] WHERE (GETDATE() = { fn NOW() })", CNN, lckLockReadOnly
With Rc
     If .DBRecordset.Recordcount <> 0 Then
        Inom = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + 1
     Else
        Inom = 1
     End If
     Select Case Len(Trim(Str(Inom)))
            Case 0: IndexAuto = "BP/" & TglSaiki & "-" & Trim(Str(Inom))
            Case 1: IndexAuto = "BP/" & TglSaiki & "-" & "0000" & Trim(Str(Inom))
            Case 2: IndexAuto = "BP/" & TglSaiki & "-" & "000" & Trim(Str(Inom))
            Case 3: IndexAuto = "BP/" & TglSaiki & "-" & "00" & Trim(Str(Inom))
            Case 4: IndexAuto = "BP/" & TglSaiki & "-" & "0" & Trim(Str(Inom))
     End Select
End With
End Function
