VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D7BB8F75-AC9E-4E80-A526-70EA20ACFD16}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmpblending 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BLENDING INSTRUCTION"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8370
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      ForeColor       =   &H80000008&
      Height          =   5370
      Left            =   0
      ScaleHeight     =   5340
      ScaleWidth      =   8355
      TabIndex        =   1
      Top             =   0
      Width           =   8385
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "id_blending"
         Height          =   375
         Index           =   0
         Left            =   2625
         TabIndex        =   14
         Tag             =   "blending"
         Top             =   450
         Width           =   1560
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "i"
         DataSource      =   "DDE"
         Height          =   375
         Index           =   10
         Left            =   4275
         TabIndex        =   13
         Tag             =   "blending"
         Top             =   2610
         Width           =   510
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "h"
         DataSource      =   "DDE"
         Height          =   375
         Index           =   9
         Left            =   4275
         TabIndex        =   12
         Tag             =   "blending"
         Top             =   2250
         Width           =   510
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "g"
         DataSource      =   "DDE"
         Height          =   375
         Index           =   8
         Left            =   2625
         TabIndex        =   11
         Tag             =   "blending"
         Top             =   2250
         Width           =   1005
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "f"
         DataSource      =   "DDE"
         Height          =   375
         Index           =   7
         Left            =   4275
         TabIndex        =   10
         Tag             =   "blending"
         Top             =   1890
         Width           =   510
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "e"
         DataSource      =   "DDE"
         Height          =   375
         Index           =   6
         Left            =   2625
         TabIndex        =   9
         Tag             =   "blending"
         Top             =   1890
         Width           =   1005
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "d"
         DataSource      =   "DDE"
         Height          =   375
         Index           =   5
         Left            =   4275
         TabIndex        =   8
         Tag             =   "blending"
         Top             =   1530
         Width           =   510
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "c"
         DataSource      =   "DDE"
         Height          =   375
         Index           =   4
         Left            =   2625
         TabIndex        =   7
         Tag             =   "blending"
         Top             =   1530
         Width           =   1005
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "b"
         DataSource      =   "DDE"
         Height          =   375
         Index           =   3
         Left            =   4275
         TabIndex        =   6
         Tag             =   "blending"
         Top             =   1170
         Width           =   510
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "lot_no"
         Height          =   375
         Index           =   1
         Left            =   2625
         TabIndex        =   5
         Tag             =   "blending"
         Top             =   810
         Width           =   1005
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "a"
         DataSource      =   "DDE"
         Height          =   375
         Index           =   2
         Left            =   2625
         TabIndex        =   4
         Tag             =   "blending"
         Top             =   1170
         Width           =   1005
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "j"
         DataSource      =   "DDE"
         Height          =   375
         Index           =   19
         Left            =   2625
         TabIndex        =   3
         Tag             =   "blending"
         Top             =   2970
         Width           =   1005
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "m"
         DataSource      =   "DDE"
         Height          =   375
         Index           =   15
         Left            =   2595
         TabIndex        =   2
         Tag             =   "blending"
         Top             =   3990
         Width           =   1005
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "k"
         DataSource      =   "DDE"
         Height          =   345
         Index           =   0
         Left            =   2595
         TabIndex        =   15
         Tag             =   "blending"
         Top             =   3315
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd hh:mm:ss"
         Format          =   58392579
         CurrentDate     =   39423
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "l"
         DataSource      =   "DDE"
         Height          =   345
         Index           =   1
         Left            =   2595
         TabIndex        =   16
         Tag             =   "blending"
         Top             =   3645
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd hh:mm:ss"
         Format          =   58392579
         CurrentDate     =   39423
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Id Blending Instruction"
         Height          =   255
         Index           =   0
         Left            =   390
         TabIndex        =   32
         Top             =   570
         Width           =   1575
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2820
         X2              =   375
         Y1              =   810
         Y2              =   810
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   4725
         X2              =   2940
         Y1              =   2970
         Y2              =   2970
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         Height          =   255
         Index           =   11
         Left            =   4845
         TabIndex        =   31
         Top             =   2700
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Powder"
         Height          =   255
         Index           =   10
         Left            =   3000
         TabIndex        =   30
         Top             =   2700
         Width           =   1050
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   4455
         X2              =   345
         Y1              =   2610
         Y2              =   2610
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   4455
         X2              =   345
         Y1              =   2250
         Y2              =   2250
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   255
         Index           =   9
         Left            =   3705
         TabIndex        =   29
         Top             =   2355
         Width           =   570
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pre Lot No                                                                                  Kg"
         Height          =   255
         Index           =   8
         Left            =   390
         TabIndex        =   28
         Top             =   2340
         Width           =   4695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   255
         Index           =   7
         Left            =   3705
         TabIndex        =   27
         Top             =   1995
         Width           =   570
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   255
         Index           =   6
         Left            =   3690
         TabIndex        =   26
         Top             =   1620
         Width           =   570
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pre Lot No                                                                                  Kg"
         Height          =   255
         Index           =   5
         Left            =   390
         TabIndex        =   25
         Top             =   1980
         Width           =   4725
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   255
         Index           =   4
         Left            =   3675
         TabIndex        =   24
         Top             =   1260
         Width           =   570
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pre Lot No                                                                                  Kg"
         Height          =   255
         Index           =   2
         Left            =   390
         TabIndex        =   23
         Top             =   1620
         Width           =   4725
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   4410
         X2              =   360
         Y1              =   1530
         Y2              =   1530
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lot No"
         Height          =   255
         Index           =   3
         Left            =   390
         TabIndex        =   22
         Top             =   900
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pre Lot No                                                                                  Kg"
         Height          =   255
         Index           =   1
         Left            =   390
         TabIndex        =   21
         Top             =   1260
         Width           =   4740
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2910
         X2              =   375
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   4455
         X2              =   345
         Y1              =   1890
         Y2              =   1890
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal dan waktu mulai"
         Height          =   255
         Index           =   20
         Left            =   345
         TabIndex        =   20
         Top             =   3720
         Width           =   2490
      End
      Begin VB.Line Line1 
         Index           =   12
         X1              =   3465
         X2              =   315
         Y1              =   3975
         Y2              =   3975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal dan waktu mulai"
         Height          =   255
         Index           =   22
         Left            =   360
         TabIndex        =   19
         Top             =   3405
         Width           =   2490
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         Height          =   255
         Index           =   21
         Left            =   345
         TabIndex        =   18
         Top             =   3060
         Width           =   645
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   3420
         X2              =   330
         Y1              =   3645
         Y2              =   3645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total powder setelah blending"
         Height          =   255
         Index           =   18
         Left            =   345
         TabIndex        =   17
         Top             =   4080
         Width           =   2220
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   3450
         X2              =   330
         Y1              =   3330
         Y2              =   3330
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   3495
         X2              =   345
         Y1              =   4350
         Y2              =   4350
      End
   End
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5385
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   1005
      BindFormTAG     =   "blending"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmpblending"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents op_shiever As frmCaller
Attribute op_shiever.VB_VarHelpID = -1
Dim rsop_shiever As New DBQuick

Private Sub DDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb

        Case tmbAddNew:
            Text1(0).Text = IndexAuto
            Text1(0).Enabled = False
    End Select

End Sub

Private Sub DDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb

        Case tmbSave:
            DDE.IsChildMemberReady = True
            simpan

        Case tmbDelete:
            DDE.PrepareDelete = "delete from blending_instruction where id_blending = '" & Text1(0).Text & "'"
    End Select

End Sub

Function simpan()
    DDE.PrepareAppend = "insert into blending_instruction (id_blending,lot_no,a,b,c,d,e,f,g,h,i,j,k,l,m) values ('" & Text1(0).Text & "',  " & " '" & DDE.GetFieldByName("lot_no") & "', '" & DDE.GetFieldByName("a") & "', '" & DDE.GetFieldByName("b") & "', '" & DDE.GetFieldByName("c") & "','" & DDE.GetFieldByName("d") & "','" & DDE.GetFieldByName("e") & "','" & DDE.GetFieldByName("f") & "','" & DDE.GetFieldByName("g") & "','" & DDE.GetFieldByName("h") & "','" & DDE.GetFieldByName("i") & "','" & DDE.GetFieldByName("j") & "', '" & Format(DTPicker1(0).value, "yyyy-MM-dd") & "', '" & Format(DTPicker1(1).value, "yyyy-MM-dd") & "', '" & DDE.GetFieldByName("m") & "') "

    DDE.PrepareUpdate = "update BLENDING_INSTRUCTION set lot_no = '" & DDE.GetFieldByName("lot_no") & "', a = '" & DDE.GetFieldByName("a") & "' , " & " b = '" & DDE.GetFieldByName("b") & "', c = '" & DDE.GetFieldByName("c") & "', " & " d = '" & DDE.GetFieldByName("d") & "', e = '" & DDE.GetFieldByName("e") & "', " & " f = '" & DDE.GetFieldByName("f") & "', g = '" & DDE.GetFieldByName("g") & "', " & " h = '" & DDE.GetFieldByName("h") & "', i = '" & DDE.GetFieldByName("i") & "', " & " j = '" & DDE.GetFieldByName("j") & "', k = '" & DDE.GetFieldByName("k") & "', " & " l = '" & DDE.GetFieldByName("l") & "', m = '" & DDE.GetFieldByName("m") & "' where id_blending = '" & Text1(0).Text & "'"
End Function

Private Sub Form_Load()

    With DDE
        Set .BindForm = Me
        .BindFormTAG = "blending"
        Set .ActiveConnection = CNN
        .PrepareQuery = "select * from blending_instruction"
    End With

    HiasForm Picture2, Me
    seting Me
End Sub

Private Function IndexAuto() As String
    Dim Rc As New DBQuick
    Dim TglSaiki As String
    Dim Inom As Long
    TglSaiki = Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2)
    Rc.DBOpen "SELECT MAX(RIGHT(id_blending, 5)) AS MaxNom FROM [BLENDING_INSTRUCTION] WHERE (GETDATE() = { fn NOW() })", CNN, lckLockReadOnly

    With Rc

        If .DBRecordset.Recordcount <> 0 Then
            Inom = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + 1
        Else
            Inom = 1
        End If

        Select Case Len(Trim(Str(Inom)))

            Case 0: IndexAuto = "BL/" & TglSaiki & "-" & Trim(Str(Inom))

            Case 1: IndexAuto = "BL/" & TglSaiki & "-" & "0000" & Trim(Str(Inom))

            Case 2: IndexAuto = "BL/" & TglSaiki & "-" & "000" & Trim(Str(Inom))

            Case 3: IndexAuto = "BL/" & TglSaiki & "-" & "00" & Trim(Str(Inom))

            Case 4: IndexAuto = "BL/" & TglSaiki & "-" & "0" & Trim(Str(Inom))
        End Select

    End With

End Function
