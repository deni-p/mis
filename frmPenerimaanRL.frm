VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmPenerimaanRL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penerimaan Rumput laut"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10665
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPenerimaanRL.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   10665
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   4005
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   1005
      BindFormTAG     =   "TTRL"
      ActiveLanguage  =   1
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4035
      Left            =   0
      ScaleHeight     =   4035
      ScaleWidth      =   10665
      TabIndex        =   22
      Top             =   0
      Width           =   10665
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "QtyReceive"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   10
         Left            =   2415
         TabIndex        =   12
         Tag             =   "TTRL"
         Top             =   2625
         Width           =   1095
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   3990
         MaskColor       =   &H000000C0&
         Picture         =   "frmPenerimaanRL.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "TTRL"
         Top             =   1208
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   4590
         MaskColor       =   &H000000C0&
         Picture         =   "frmPenerimaanRL.frx":D0A4
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "TTRL"
         Top             =   848
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   6030
         MaskColor       =   &H000000C0&
         Picture         =   "frmPenerimaanRL.frx":138F6
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "TTRL"
         Top             =   488
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAAF6F&
         Caption         =   " Kondisi "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6480
         TabIndex        =   24
         Top             =   2760
         Width           =   3015
         Begin VB.OptionButton kondisi 
            BackColor       =   &H00EAAF6F&
            Caption         =   "Kering"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   1560
            TabIndex        =   21
            Tag             =   "TTRL"
            Top             =   330
            Width           =   855
         End
         Begin VB.OptionButton kondisi 
            BackColor       =   &H00EAAF6F&
            Caption         =   "Basah"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Tag             =   "TTRL"
            Top             =   330
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAAF6F&
         Caption         =   " Jenis "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   6480
         TabIndex        =   23
         Top             =   1560
         Width           =   3975
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   9
            Left            =   1365
            TabIndex        =   19
            Top             =   600
            Width           =   2415
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00EAAF6F&
            Caption         =   "Lainnya"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   18
            Tag             =   "TTRL"
            Top             =   668
            Width           =   960
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00EAAF6F&
            Caption         =   "Gracilaria"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Tag             =   "TTRL"
            Top             =   300
            Width           =   2055
         End
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "NoPol"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   8
         Left            =   2415
         TabIndex        =   14
         Tag             =   "TTRL"
         Top             =   2978
         Width           =   2790
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "SakAktual"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   6
         Left            =   4110
         TabIndex        =   13
         Tag             =   "TTRL"
         Top             =   2625
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "SakSJ"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   5
         Left            =   4110
         TabIndex        =   11
         Tag             =   "TTRL"
         Top             =   2265
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "beratSJ"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   4
         Left            =   2415
         TabIndex        =   10
         Tag             =   "TTRL"
         Top             =   2265
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "noItem"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   3
         Left            =   2415
         TabIndex        =   7
         Tag             =   "TTRL"
         Top             =   1203
         Width           =   1575
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "NoPO"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   2
         Left            =   2415
         TabIndex        =   5
         Tag             =   "TTRL"
         Top             =   842
         Width           =   2175
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "companyName"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   1
         Left            =   2415
         TabIndex        =   3
         Tag             =   "TTRL"
         Top             =   481
         Width           =   3615
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "ID"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   2415
         TabIndex        =   1
         Tag             =   "TTRL"
         Top             =   120
         Width           =   2415
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "warehouse"
         DataSource      =   "DDE"
         Height          =   330
         Index           =   7
         Left            =   2415
         TabIndex        =   16
         Tag             =   "TTRL"
         Top             =   3345
         Width           =   2790
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   5205
         MaskColor       =   &H000000C0&
         Picture         =   "frmPenerimaanRL.frx":1A148
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "TTRL"
         Top             =   3353
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "tgl"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   0
         Left            =   2415
         TabIndex        =   8
         Tag             =   "TTRL"
         Top             =   1564
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   64684035
         CurrentDate     =   39365
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "tgl"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   1
         Left            =   2415
         TabIndex        =   9
         Tag             =   "TTRL"
         Top             =   1910
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Format          =   64684034
         CurrentDate     =   39365
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   13
         Left            =   4560
         TabIndex        =   38
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty PO :"
         Height          =   255
         Index           =   6
         Left            =   9285
         TabIndex        =   37
         Top             =   135
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   8
         X1              =   2400
         X2              =   360
         Y1              =   3293
         Y2              =   3293
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   7
         X1              =   2400
         X2              =   360
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   6
         X1              =   2400
         X2              =   360
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   5
         X1              =   2400
         X2              =   360
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   4
         X1              =   2400
         X2              =   360
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   3
         X1              =   2400
         X2              =   360
         Y1              =   1518
         Y2              =   1518
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   2
         X1              =   2400
         X2              =   360
         Y1              =   1157
         Y2              =   1157
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   1
         X1              =   2400
         X2              =   360
         Y1              =   796
         Y2              =   796
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   0
         X1              =   2400
         X2              =   360
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Item"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   12
         Left            =   420
         TabIndex        =   36
         Top             =   1275
         Width           =   570
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Pol Kendaraan"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   11
         Left            =   420
         TabIndex        =   35
         Top             =   3045
         Width           =   1335
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Kg                                 Sak"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   9
         Left            =   3600
         TabIndex        =   34
         Top             =   2333
         Width           =   1920
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty Aktual"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   8
         Left            =   420
         TabIndex        =   33
         Top             =   2700
         Width           =   765
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty di Surat Jalan"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   420
         TabIndex        =   32
         Top             =   2340
         Width           =   1290
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   420
         TabIndex        =   31
         Top             =   1965
         Width           =   285
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Datang"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   420
         TabIndex        =   30
         Top             =   1620
         Width           =   1140
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. RPB"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   420
         TabIndex        =   29
         Top             =   915
         Width           =   585
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   28
         Top             =   555
         Width           =   570
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   27
         Top             =   195
         Width           =   465
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gudang"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   420
         TabIndex        =   26
         Top             =   3420
         Width           =   555
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   9
         X1              =   2400
         X2              =   360
         Y1              =   3660
         Y2              =   3660
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kg                                 Sak"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   10
         Left            =   3600
         TabIndex        =   25
         Top             =   2693
         Width           =   1920
      End
   End
End
Attribute VB_Name = "frmPenerimaanRL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MaxQty As Double
Dim RcQty As New DBQuick
Dim RcPartner As New DBQuick
Dim IDGen As New IDGenerator
Dim newData As Boolean
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1

Private Sub InsertTransData()
Dim lJenis As String
Dim lKondisi As String
   lJenis = IIf(Option1(0).Value = True, "Gracilaria", Trim(txt(9).Text))
   lKondisi = IIf(kondisi(0).Value, "Basah", "Kering")

With DDE
      SendDataToServer " insert into [detail TransData] (TransID,referense," & _
                                              "NoItem,DateTrans," & _
                                              "jenisRL," & _
                                              "kondisi," & _
                                              "Qty_in," & _
                                              "refNotes," & _
                                              "Qty_Receive," & _
                                              "sak," & _
                                              "status) " & _
                         " values ('" & .GetFieldByName("ID") & _
                                "','" & .GetFieldByName("NoPO") & _
                                "','" & .GetFieldByName("NoItem") & _
                                "','" & Format(DTPicker1(0).Value, "yyyy-MM-dd") & " " & Format(DTPicker1(1).Value, "hh:mm:ss") & _
                                "','" & lJenis & _
                                "','" & lKondisi & _
                                "', " & .GetFieldByName("BeratSJ") & _
                                " , " & .GetFieldByName("sakSJ") & _
                                " , " & .GetFieldByName("QtyReceive") & _
                                " , " & .GetFieldByName("SakAktual") & _
                                ", 0)"

End With
End Sub

Private Sub UpdateTransData()
Dim lJenis As String
Dim lKondisi As String
   lJenis = IIf(Option1(0).Value = True, "Gracilaria", Trim(txt(9).Text))
   lKondisi = IIf(kondisi(0).Value, "Basah", "Kering")

With DDE
    
    SendDataToServer " update [detail TransData] set Referense='" & .GetFieldByName("NoPo") & _
                                               "',NoItem='" & .GetFieldByName("NoItem") & _
                                               "',DateTrans='" & Format(DTPicker1(0).Value, "yyyy-MM-dd") & " " & Format(DTPicker1(1).Value, "hh:mm:ss") & _
                                               "',jenisRL='" & lJenis & _
                                               "',kondisi='" & lKondisi & _
                                               "',Qty_in=" & .GetFieldByName("beratSJ") & _
                                               ",refNotes=" & .GetFieldByName("sakSJ") & _
                                               ",Qty_receive=" & .GetFieldByName("QtyReceive") & _
                                               ",sak=" & .GetFieldByName("sakAktual") & _
                      " where TransID='" & .GetFieldByName("ID") & "' and noItem='" & .GetFieldByName("NoItem") & "'"

End With
End Sub


Private Sub DeleteTransData()
   SendDataToServer "delete from TransData where TransID='" & txt(0).Text & "'"
End Sub


Private Sub cmdLink_Click(Index As Integer)
   Select Case Index
      Case 0:
         OpenPartner Index
      Case 1:
         If Not IsNull(DDE.GetFieldByName("PartnerID")) Then
            OpenPartner Index, DDE.GetFieldByName("PartnerID")
         End If
      Case 2:
         If Not IsNull(DDE.GetFieldByName("NoPo")) Then
            OpenPartner Index, DDE.GetFieldByName("NoPo")
         End If
      Case 3:
         OpenPartner Index
   End Select
   
End Sub

Private Sub DDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
   CmdLink(0).Enabled = False
   CmdLink(1).Enabled = False
   CmdLink(2).Enabled = False
   CmdLink(3).Enabled = False
   
   Select Case AdReasonActiveDb
      Case tmbAddNew:
         DDE.GetFieldByName("ID") = IDGen.GetID("RR")
         DTPicker1(0).Value = Now
         DTPicker1(1).Value = Now
         DDE.GetFieldByName("tgl") = DTPicker1(0).Value
         CmdLink(0).Enabled = True
         CmdLink(1).Enabled = True
         CmdLink(2).Enabled = True
         CmdLink(3).Enabled = True
      Case tmbEdit:
         CmdLink(0).Enabled = True
         CmdLink(1).Enabled = True
         CmdLink(2).Enabled = True
         CmdLink(3).Enabled = True
      Case tmbSave:
         If Not DDE.CancelTrans Then
            If DDE.IsChildMemberReady Then
               'save to table [detail PO] untuk merubah nilai QtyReceive
               If newData Then
                  SendDataToServer "update [detail PO] set QtyReceive = " & Val(txt(10).Text) + Val(RcQty.Fields("QtyReceive")) & ",StatusTrans=2,DNID='" & DDE.GetFieldByName("ID") & "',QtyTemp = " & Val(RcQty.Fields("QtyTemp")) - Val(txt(10).Text) & " where PurchaseID='" & DDE.GetFieldByName("NoPO") & "' and noItem='" & DDE.GetFieldByName("NoItem") & "'"
                  InsertTransData
                  InsertStock
               Else
                  SendDataToServer "update [detail PO] set QtyReceive = " & txt(10).Text & ",StatusTrans=2,DNID='" & DDE.GetFieldByName("ID") & "',QtyTemp=" & Val(RcQty.Fields("QtyTemp")) - Val(txt(10).Text) & " where PurchaseID='" & DDE.GetFieldByName("NoPO") & "' and noItem='" & DDE.GetFieldByName("NoItem") & "'"
                  UpdateTransData
                  UpdateStock
               End If
               
               'Update status [PO Order] = 1 / closed
               If Val(RcQty.Fields("QtyTemp")) - Val(txt(10).Text) = 0 Then
                  SendDataToServer "update [PO Order] set status=1 where PurchaseID='" & txt(2).Text & "'"
               End If
               
            End If
         End If
      
      Case tmbDelete:
         If Not DDE.CancelTrans Then
            CancelDetailPO
            DeleteTransData
         End If
      
      Case tmbPrint:
         frmSelection.ReportFile = "TandaTerimaRL.rpt"
         frmSelection.ID = txt(0).Text '& txt(3).Text
         If Option1(0).Value Then
            frmSelection.JenisRL = "Gracilaria"
         Else
            frmSelection.JenisRL = txt(9).Text
         End If
         frmSelection.Suppplier = txt(1).Text
         frmSelection.BeratRL = txt(10).Text & " Kg"
         frmSelection.sql = "select * from QueryPenerimaanRL where id='" & txt(0).Text & "'"
         frmSelection.Show vbModal
   End Select
End Sub

Private Sub InsertStock()
   Dim rsCons As New DBQuick
   Dim unitKoversi As Double
   rsCons.DBOpen "select UOMKonversi from Inventory where noItem ='" & DDE.GetFieldByName("NoItem") & "'", CNN, lckLockReadOnly
   If rsCons.Recordcount > 0 Then
      unitKoversi = Val(rsCons.DBRecordset.Fields(0)) * Val(DDE.GetFieldByName("QtyReceive"))
   Else
      unitKoversi = Val(DDE.GetFieldByName("QtyReceive"))
   End If
   SendDataToServer "Insert into [inventory tabel] (noIdx,noItem,Qty_in,Qty_out,refTrans,DateTrans,StockTmp,LockFIFO,TypeTrans,sl_no) values (newID(),'" & _
                                                    DDE.GetFieldByName("NoItem") & _
                                                    "'," & FQty(unitKoversi) & _
                                                    ",0,'" & DDE.GetFieldByName("ID") & _
                                                    "','" & Format(DTPicker1(0).Value, "yyyy-MM-dd") & _
                                                    "', " & FQty(unitKoversi) & _
                                                    ",0,'RL','" & txt(0) & "')"
 Set rsCons = Nothing
End Sub

Private Sub UpdateStock()
   Dim rsBalance As New DBQuick
   Dim SisaStock As Double
   Dim rsCons As New DBQuick
   Dim unitKoversi As Double
   rsCons.DBOpen "select UOMKonversi from Inventory where noItem ='" & DDE.GetFieldByName("NoItem") & "'", CNN, lckLockReadOnly
   If rsCons.Recordcount > 0 Then
      unitKoversi = Val(rsCons.DBRecordset.Fields(0))
   Else
      unitKoversi = 1
   End If
   
   rsBalance.DBOpen "select Qty_out,StockTmp from [inventory Tabel] where noItem='" & DDE.GetFieldByName("noItem") & "' and refTrans='" & DDE.GetFieldByName("ID") & "' and LokasiGdg = '" & txt(7).Text & "'", CNN, lckLockReadOnly
   
   If rsBalance.DBRecordset.Recordcount > 0 Then
      If Val(DDE.GetFieldByName("QtyReceive")) <= (Val(rsBalance.DBRecordset.Fields("Qty_out")) / unitKoversi) Then
         SisaStock = Val(DDE.GetFieldByName("QtyReceive")) - Val(rsBalance.DBRecordset.Fields("Qty_out"))
         SendDataToServer "update [inventory tabel] set Qty_in=" & DDE.GetFieldByName("QtyReceive") & _
                                                         ",QtyTemp=" & SisaStock & _
                                                         ",LockFifo = " & IIf(SisaStock = 0, "1", "0") & _
                          " where noItem='" & DDE.GetFieldByName("noItem") & "' and  refTrans='" & DDE.GetFieldByName("ID") & "' and lokasiGdg ='" & txt(7).Text & "'"
      End If
   End If
   Set rsBalance = Nothing
End Sub

Private Sub DDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo xErr
   lbl(13).Caption = DDE.GetFieldByName("QtyPO")
   If UCase(Trim(DDE.GetFieldByName("jenisRL"))) = "GRACILARIA" Then
      Option1(0).Value = True
      Option1(1).Value = False
      txt(9).Text = ""
   Else
      Option1(0).Value = False
      Option1(1).Value = True
      txt(9).Text = DDE.GetFieldByName("jenisRL")
   End If
   
   If UCase(Trim(DDE.GetFieldByName("Kondisi"))) = "BASAH" Then
      kondisi(0).Value = True
      kondisi(1).Value = False
   Else
      kondisi(0).Value = False
      kondisi(1).Value = True
   End If
Exit Sub
xErr:
   Option1(0).Value = True
   Option1(1).Value = False
   txt(9).Text = "xxxxx"
   kondisi(0).Value = True
   kondisi(1).Value = False
   Err.Clear
End Sub


Private Sub CancelDetailPO()
   SendDataToServer "update [detail PO] set DNID= null, QtyReceive = 0 where PurchaseID='" & DDE.GetFieldByName("NoPO") & "' and noItem='" & DDE.GetFieldByName("NoItem") & "'"
End Sub

Private Sub PrepareSQL()
   
   With DDE
        .PrepareAppend = "insert into TransData(TransID," & _
                                          "EmpID," & _
                                          "DateTrans," & _
                                          "purchaseID," & _
                                          "PartnerID," & _
                                          "[No Pol]," & _
                                          "TypeTrans,warehouse)" & _
                   " values ('" & txt(0).Text & _
                          "','" & MainMenu.StatusBar1.Panels(1).Text & _
                          "','" & Format(DTPicker1(0).Value, "yyyy-mm-dd") & _
                          "','" & txt(2).Text & _
                          "','" & DDE.GetFieldByName("PartnerID") & _
                          "','" & txt(8).Text & _
                          "','AP','" & txt(7).Text & "')"
      
      
         .PrepareUpdate = "update TransData Set EmpID ='" & MainMenu.StatusBar1.Panels(1).Text & _
                                   "',DateTrans ='" & Format(DTPicker1(0).Value, "yyyy-MM-dd") & _
                                   "',PurchaseID='" & txt(2).Text & _
                                   "',PartnerID='" & DDE.GetFieldByName("PartnerID") & _
                                   "',[No Pol]='" & txt(8).Text & _
                                   "',TypeTrans='AP' " & _
                     " where TransID='" & txt(0).Text & "'"

                                
        

      .PrepareDelete = " delete from transData where TransID='" & .GetFieldByName("ID") & "'"
   
   
   End With
End Sub

Private Sub DDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

   Select Case AdReasonActiveDb
      Case tmbAddNew:
         newData = True
      Case tmbEdit:
         newData = False
      Case tmbSave:
         If DDE.CheckEmptyControl = False Then
            DDE.IsChildMemberReady = True
            If newData Then
               If Not IDGen.IsValidID Then
                  DDE.GetFieldByName("ID") = IDGen.GetID("RR")
               End If
               
               If DDE.GetFieldByName("QtyReceive") > RcQty.Fields("MaxQty") Then
                  MessageBox "Max Qty yang diperbolehkan Adalah " & RcQty.Fields("MaxQty"), "Informasi", msgOkOnly, msgInfo
                  DDE.CancelTrans = True
               End If
            Else
               If DDE.GetFieldByName("QtyReceive") > DDE.GetFieldByName("QtyPO") Then
                  MessageBox "Max Qty yang diperbolehkan Adalah " & DDE.GetFieldByName("QtyPO"), "Informasi", msgOkOnly, msgInfo
                  DDE.CancelTrans = True
               End If
            End If
            PrepareSQL
         Else
            DDE.IsChildMemberReady = False
         End If
      Case tmbDelete:
         PrepareSQL
   End Select
End Sub



Private Sub OpenPartner(ByVal Index As Integer, Optional Params As String)
On Error GoTo Hell:
If Params = "" Then Params = "xxxxxxxxxxxxxx"
Select Case Index
       Case 0:  'supplier
            RcPartner.DBOpen "SELECT PartnerID AS [Partner ID], CompanyName AS Perusahaan, Address AS Alamat, " & _
                                      "City AS Kota, PostalCode AS [Kode Pos], Country AS Negara, Phone AS Telp, " & _
                                      "[Due Date Calculation] AS termPayment, code " & _
                                      " FROM  supplier_blanked", CNN, lckLockReadOnly
                                      
                                '      Debug.Print "SELECT PartnerDB.PartnerID AS [Partner ID], PartnerDB.CompanyName AS Perusahaan, PartnerDB.Address AS Alamat, " & _
                                      "PartnerDB.City AS Kota, PartnerDB.PostalCode AS [Kode Pos], PartnerDB.Country AS Negara, PartnerDB.Phone AS Telp, " & _
                                      "TermPayment.[Due Date Calculation] AS termPayment, termPayment.code " & _
                                      " FROM  PartnerDB LEFT OUTER JOIN " & _
                                      " TermPayment ON PartnerDB.Term_code = TermPayment.Code " & _
                                      " WHERE   (PartnerDB.PartnerType = N'SUPPLIER') " & _
                                      " ORDER BY PartnerDB.CompanyName"
       Case 1: 'PO
            RcPartner.DBOpen "select PurchaseID as NoPO, DatePurchase as Tanggal from [PO Order] where status = 2 and typeTRans='PO' and partnerID ='" & Params & "' order by datePurchase desc ", CNN, lckLockReadOnly
       Case 2: 'Item
            RcPartner.DBOpen "SELECT [detail PO].noItem, Inventory.ItemName as Nama, Inventory.UOM as Unit, [detail PO].QtyPO as QTY  FROM [Detail PO] INNER JOIN Inventory ON [Detail PO].NoItem = Inventory.NoItem where [detail PO].purchaseID ='" & Params & "' and (([detail PO].QtyPO - [detail PO].QtyReceive) > 0) and [detail PO].noItem like 'BB%' ", CNN, lckLockReadOnly
       Case 3: 'Gudang
            RcPartner.DBOpen "select warehouse as Kode,[Warehouse Name] as Gudang ,locations as lokasi from WareHouse", CNN, lckLockReadOnly
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 0:
            mCall.FromTagActive = "Daftar Supplier"
            mCall.CaptionLink = "Supplier"
          Case 1:
            mCall.FromTagActive = "Daftar PO"
            mCall.CaptionLink = "Daftar PO"
          Case 2:
            mCall.FromTagActive = "Item PO"
            mCall.CaptionLink = "Item PO"
          Case 3:
            mCall.FromTagActive = "Gudang"
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me

Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgCrtical
End If
'
Exit Sub
Hell:
    Err.Clear
End Sub


Private Sub Form_Load()
   'HiasForm Picture1, Me
   HiasFormManTell Picture2, Me
   Set mCall = New frmCaller
   Set DDE.ActiveConnection = CNN
   Set DDE.BindForm = Me
   'DDE.PrepareQuery = "SELECT  PenerimaanRL.*, PartnerDB.CompanyName,partnerDB.partnerID, [Detail PO].QTYPO, [Detail PO].QTYReceive, [detail PO].QtyReceive as QtyReceived " & _
   '                   "FROM [PO Order] INNER JOIN " & _
   '                   "PartnerDB ON [PO Order].PartnerID = PartnerDB.PartnerID INNER JOIN " & _
   '                   "PenerimaanRL ON [PO Order].PurchaseID = PenerimaanRL.NoPO INNER JOIN " & _
   '                   "[Detail PO] ON PenerimaanRL.NoPO = [Detail PO].PurchaseID AND PenerimaanRL.NoItem = [Detail PO].NoItem"
   DDE.PrepareQuery = "select * from QueryPenerimaanRL order by tgl desc"
   DDE.SetPermissions = aksess.MayDo("Tanda Terima RL")
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
   Select Case TagForm
      Case "Daftar Supplier":
         DDE.GetFieldByName("PartnerID") = mCall.GetFieldByName("Partner ID")
         DDE.GetFieldByName("companyName") = mCall.GetFieldByName("Perusahaan")
         If newData Then
            IDGen.ExtParameter = Left(mCall.GetFieldByName("Perusahaan"), 4) & "-" & Left(mCall.GetFieldByName("Kota"), 4)
            txt(0).Text = IDGen.GetID("RR")
            DDE.GetFieldByName("ID") = txt(0).Text
         End If
      Case "Daftar PO":
         DDE.GetFieldByName("NoPO") = mCall.GetFieldByName("NoPO")
         
      Case "Item PO":
         DDE.GetFieldByName("NoItem") = mCall.GetFieldByName("NoItem")
         RcQty.DBOpen "select QtyPO,QtyReceive,(QtyPO - QtyReceive) as maxQty, QtyTemp  from [detail PO] where purchaseID='" & DDE.GetFieldByName("NoPo") & "' and NoItem='" & DDE.GetFieldByName("NoItem") & "'", CNN, lckLockReadOnly
         DDE.GetFieldByName("QtyPO") = RcQty.Fields("QtyPO")
         DDE.GetFieldByName("QtyReceive") = 0
         DDE.GetFieldByName("BeratSJ") = mCall.GetFieldByName("QTY")
         DDE.GetFieldByName("sakSJ") = 1
         DDE.GetFieldByName("sakAktual") = 1
         
      Case "Gudang"
         txt(7).Text = mCall.GetFieldByName(0)
         
   End Select
End Sub

