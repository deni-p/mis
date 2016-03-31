VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{E224F788-0398-4D72-B72C-F9D023C39E0D}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmMilling 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MILLING POWDER"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9555
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMilling.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4860
      Left            =   0
      ScaleHeight     =   4860
      ScaleWidth      =   9555
      TabIndex        =   1
      Top             =   0
      Width           =   9555
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "grup"
         DataSource      =   "DDE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2010
         TabIndex        =   10
         Tag             =   "miling"
         Top             =   555
         Width           =   2370
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "pre_lot_powder"
         DataSource      =   "DDE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2010
         TabIndex        =   9
         Tag             =   "miling"
         Top             =   210
         Width           =   1950
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "powder_1"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   7170
         TabIndex        =   7
         Tag             =   "miling"
         Top             =   555
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "powder_2"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   7170
         TabIndex        =   6
         Tag             =   "miling"
         Top             =   900
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "total_powder"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   7170
         TabIndex        =   5
         Tag             =   "miling"
         Top             =   1260
         Width           =   645
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3990
         MaskColor       =   &H000000C0&
         Picture         =   "frmMilling.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "powder"
         Top             =   210
         UseMaskColor    =   -1  'True
         Width           =   390
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "id_mixing_chips"
         DataSource      =   "DDE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   180
         TabIndex        =   2
         Tag             =   "miling"
         Top             =   1275
         Visible         =   0   'False
         Width           =   1710
      End
      Begin MSDataGridLib.DataGrid dgdetail 
         Height          =   2835
         Left            =   165
         TabIndex        =   4
         Top             =   1740
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   5001
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   15
         BeginProperty Column00 
            DataField       =   "waktu"
            Caption         =   "Waktu"
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
            DataField       =   "a"
            Caption         =   "A > 80"
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
            DataField       =   "b"
            Caption         =   "B 80"
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
         BeginProperty Column03 
            DataField       =   "c"
            Caption         =   "C 100"
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
         BeginProperty Column04 
            DataField       =   "total_1"
            Caption         =   "Total"
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
         BeginProperty Column05 
            DataField       =   "d"
            Caption         =   "% A > 80"
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
         BeginProperty Column06 
            DataField       =   "e"
            Caption         =   "% B 80"
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
         BeginProperty Column07 
            DataField       =   "f"
            Caption         =   "% C 100"
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
         BeginProperty Column08 
            DataField       =   "total_2"
            Caption         =   "Total"
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
         BeginProperty Column09 
            DataField       =   "g"
            Caption         =   "A > 150"
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
         BeginProperty Column10 
            DataField       =   "h"
            Caption         =   "B 150"
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
         BeginProperty Column11 
            DataField       =   "Total_3"
            Caption         =   "Total"
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
         BeginProperty Column12 
            DataField       =   "i"
            Caption         =   "% A > 150"
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
         BeginProperty Column13 
            DataField       =   "j"
            Caption         =   "% B 150"
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
         BeginProperty Column14 
            DataField       =   "total_4"
            Caption         =   "Total"
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
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
            EndProperty
            BeginProperty Column13 
            EndProperty
            BeginProperty Column14 
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker tgl 
         Height          =   315
         Index           =   0
         Left            =   1995
         TabIndex        =   8
         Tag             =   "miling"
         Top             =   900
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MMMM-yyyy"
         Format          =   64094211
         CurrentDate     =   39427
      End
      Begin MSComCtl2.DTPicker tgl 
         Height          =   315
         Index           =   2
         Left            =   7170
         TabIndex        =   11
         Tag             =   "miling"
         Top             =   210
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MMMM-yyyy"
         Format          =   64094211
         CurrentDate     =   39427
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   5175
         X2              =   7500
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   2655
         X2              =   195
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2730
         X2              =   195
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2640
         X2              =   195
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   7650
         X2              =   5175
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal && Waktu Selesai"
         Height          =   255
         Index           =   4
         Left            =   5175
         TabIndex        =   15
         Top             =   255
         Width           =   2550
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Powder                                       Kg"
         Height          =   255
         Index           =   5
         Left            =   5175
         TabIndex        =   14
         Top             =   1320
         Width           =   2940
      End
      Begin VB.Line Line2 
         Index           =   4
         X1              =   5175
         X2              =   7500
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line2 
         Index           =   3
         X1              =   5175
         X2              =   7275
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Powder mesh  150                      Kg"
         Height          =   255
         Index           =   17
         Left            =   5175
         TabIndex        =   13
         Top             =   975
         Width           =   3045
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Powder mesh > 150                    Kg"
         Height          =   255
         Index           =   16
         Left            =   5175
         TabIndex        =   12
         Top             =   615
         Width           =   3000
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal && Waktu Mulai"
         Height          =   255
         Index           =   2
         Left            =   195
         TabIndex        =   18
         Top             =   960
         Width           =   1710
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         Height          =   255
         Index           =   1
         Left            =   195
         TabIndex        =   17
         Top             =   585
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pre Lot Powder No"
         Height          =   255
         Index           =   0
         Left            =   195
         TabIndex        =   16
         Top             =   255
         Width           =   1575
      End
   End
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   4860
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   1005
      BindFormTAG     =   "mixing"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmMilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mixing As frmCaller
Attribute mixing.VB_VarHelpID = -1
Dim rsmixing As New DBQuick
Dim tabel As String

Function grid()
     DGDETAIL.AllowAddNew = True
     DGDETAIL.AllowDelete = True
     DGDETAIL.AllowUpdate = True
End Function

Private Sub cmdLink_Click()
rsmixing.DBOpen "select * from mixing_chips ", CNN
Set mixing = New frmCaller
Set mixing.FormData = rsmixing.DBRecordset
mixing.FromTagActive = "MIXING CHIPS"
mixing.CaptionLink = "MIXING CHIPS"
End Sub

Private Sub DDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
Case tmbAddNew:
     cmdLink.Enabled = True

     Text1(5).Enabled = False
     grid
Case tmbEdit:
     grid
End Select

End Sub
Function simpan()
DDE.PrepareAppend = "insert into " & tabel & " (id_mixing_chips, pre_lot_powder, grup, tgl_mulai, waktu_mulai, tgl_selesai, waktu_selesai, powder_1, powder_2, total_powder) values ('" & Text1(5).Text & "','" & Text1(0).Text & "', '" & Text1(1).Text & "', " & _
                        " '" & Format(tgl(0).value, "yyyy-MM-dd") & "','" & Format(tgl(1).value, "hh:mm:ss") & "', '" & Format(tgl(2).value, "yyyy-MM-dd") & "','" & Format(tgl(3).value, "hh:mm:ss") & "', " & _
                        " '" & Text1(2).Text & "', '" & Text1(3).Text & "', '" & Text1(4).Text & "')"
                        
DDE.PrepareUpdate = " update " & tabel & " set  pre_lot_powder = '" & Text1(0).Text & "',grup = '" & Text1(1).Text & "', tgl_mulai = '" & Format(tgl(0).value, "yyyy-MM-dd") & "', waktu_mulai = '" & Format(tgl(1).value, "hh:mm:ss") & "', tgl_selesai = '" & Format(tgl(2).value, "yyyy-MM-dd") & "', " & _
                    " waktu_selesai = '" & Format(tgl(3).value, "hh:mm:ss") & "', powder_1 = '" & Text1(2).Text & "', powder_2 = '" & Text1(3).Text & "', total_powder = '" & Text1(4).Text & "' where id_mixing_chips = '" & DDE.GetFieldByName("id_mixing_chips") & "'"

End Function


Private Sub DDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim rsdetail As New DBQuick
rsdetail.DBOpen "select * from view_milling where id_mixing_chips = '" & DDE.GetFieldByName("id_mixing_chips") & "' ", CNN
Set DDE.ChildRecordset = rsdetail.DBRecordset.Clone(adLockBatchOptimistic)
Set DGDETAIL.DataSource = DDE.ChildRecordset
End Sub

Private Sub DDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

Select Case AdReasonActiveDb
Case tmbSave:
    DDE.IsChildMemberReady = True
    simpan
    simpan_detail
Case tmbDelete:
    DDE.PrepareDelete = "delete " & tabel & " where id_mixing_chips = '" & DDE.GetFieldByName("id_mixing_chips") & "'"
End Select
End Sub



Private Sub DGDETAIL_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
DDE.ChildRecordset.AddNew
End If
End Sub

Private Sub Form_Load()
tabel = "MILLING_POWDER_HEADER"
With DDE
Set .BindForm = Me
    .BindFormTAG = "miling"
Set .ActiveConnection = CNN
    .PrepareQuery = " select * from " & tabel & " "
End With
HiasForm Picture2, Me
seting Me
End Sub

Function simpan_detail()
With DDE.ChildRecordset
   If .Recordcount <> 0 Then
       .MoveFirst
         If SendDataToServer(" delete from [MILLING_POWDER_detail] where (id_mixing_chips= '" & Text1(0).Text & "')") = True Then
         Do
           If .EOF = True Then Exit Do
           SendDataToServer "insert into MILLING_POWDER_detail (id_mixing_chips,waktu,a,b,c,total_1,d,e,f,total_2,g,h,total_3,i,j,total_4)  " & _
           " values ('" & Text1(5).Text & "','" & .Fields("waktu") & "','" & .Fields("a") & "', " & _
           " '" & .Fields("b") & "', " & _
           " '" & .Fields("c") & "', " & _
           " '" & .Fields("total_1") & "','" & .Fields("d") & "', " & _
           " '" & .Fields("e") & "','" & .Fields("f") & "', " & _
           " '" & .Fields("total_2") & "','" & .Fields("g") & "', " & _
           " '" & .Fields("h") & "','" & .Fields("total_3") & "', " & _
           " '" & .Fields("i") & "','" & .Fields("j") & "', " & _
           " '" & .Fields("total_4") & "')"
          .MoveNext
        Loop
        End If
        .MoveLast
        DGDETAIL.Refresh
        End If
    End With
End Function

Private Sub mixing_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Text1(0).Text = rsmixing.DBRecordset.Fields("pre_lot_chip_no")
Text1(5).Text = rsmixing.DBRecordset.Fields("id_mixing_chips")
End Sub

