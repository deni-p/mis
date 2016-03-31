VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{D7BB8F75-AC9E-4E80-A526-70EA20ACFD16}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FRMSETUP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SETUP ENTRY DATA"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9045
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   9045
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "id_procedure"
      DataSource      =   "dde"
      Enabled         =   0   'False
      Height          =   345
      Index           =   2
      Left            =   1305
      TabIndex        =   7
      Tag             =   "ext"
      Top             =   540
      Width           =   1905
   End
   Begin MSDataGridLib.DataGrid DGDETAIL 
      Bindings        =   "FRMSETUP.frx":0000
      Height          =   2295
      Left            =   240
      TabIndex        =   6
      Tag             =   "ext"
      Top             =   1575
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   4048
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "ID_ANALISA"
         Caption         =   "ID_ANALISA"
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
         DataField       =   "NAMA_ANALISA"
         Caption         =   "NAMA_ANALISA"
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
            ColumnWidth     =   4004,788
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "id_form"
      DataSource      =   "DDE"
      Enabled         =   0   'False
      Height          =   345
      Index           =   1
      Left            =   1305
      TabIndex        =   4
      Tag             =   "ext"
      Top             =   210
      Width           =   1905
   End
   Begin VB.CommandButton command1 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4725
      MaskColor       =   &H000000C0&
      Picture         =   "FRMSETUP.frx":0012
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "minta_sampel"
      Top             =   885
      UseMaskColor    =   -1  'True
      Width           =   350
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "nama_procedure"
      DataSource      =   "DDE"
      Enabled         =   0   'False
      Height          =   345
      Index           =   0
      Left            =   1305
      TabIndex        =   1
      Tag             =   "ext"
      Top             =   870
      Width           =   3330
   End
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Tag             =   "ext"
      Top             =   4020
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   1005
      BindFormTAG     =   "ext"
      InitControlSet  =   1
   End
   Begin VB.Label Label1 
      Caption         =   "ID PRO"
      Height          =   255
      Index           =   2
      Left            =   225
      TabIndex        =   8
      Top             =   630
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "CONFIG pRO"
      Height          =   255
      Index           =   1
      Left            =   225
      TabIndex        =   5
      Top             =   270
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Procedure"
      Height          =   255
      Index           =   0
      Left            =   270
      TabIndex        =   2
      Top             =   960
      Width           =   1365
   End
End
Attribute VB_Name = "FRMSETUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents Manalisa As frmCaller
Attribute Manalisa.VB_VarHelpID = -1

Private WithEvents Mprocedure As frmCaller
Attribute Mprocedure.VB_VarHelpID = -1
Dim rsprocedure As New DBQuick
Dim rsanalisa As New DBQuick
Dim rsDetail As New DBQuick

Private Sub Command1_Click()
    rsprocedure.DBOpen "select * from FRM_PROCEDURE", CNN
    Set Mprocedure = New frmCaller
    Set Mprocedure.FormData = rsprocedure.DBRecordset
    Mprocedure.FromTagActive = "PROCEDURE"
    Mprocedure.CaptionLink = "PROCEDURE"
End Sub

Private Sub DDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
                             ByVal pError As ADODB.Error, _
                             adStatus As ADODB.EventStatusEnum, _
                             ByVal pRecordset As ADODB.Recordset)

    Set rsDetail = New DBQuick
    rsDetail.DBOpen "select * from view_analisa where id_procedure = '" & DDE.GetFieldByName("id_procedure") & "'", CNN
    Set DDE.ChildRecordset = rsDetail.DBRecordset
    Set DgDetail.DataSource = DDE.ChildRecordset

End Sub

Private Sub DDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb

        Case tmbAddNew
            Command1.Enabled = True

        Case tmbSave:
            DDE.IsChildMemberReady = True
            simpan
            simpan_detail

        Case tmbDetail:
            rsanalisa.DBOpen "select * from FRM_ANALISA", CNN
            Set Manalisa = New frmCaller
            Set Manalisa.FormData = rsanalisa.DBRecordset
            Manalisa.FromTagActive = "ANALISA"
            Manalisa.CaptionLink = "ANALISA"
    End Select

End Sub

Function simpan()
    DDE.PrepareAppend = "insert into FRM_HEADER (id_form, id_procedure) values ('" & DDE.GetFieldByName("id_form") & "', '" & DDE.GetFieldByName("id_procedure") & "')"
End Function

Function simpan_detail()

    With DDE.ChildRecordset

        If .Recordcount <> 0 Then
            .MoveFirst

            If SendDataToServer(" delete from [FRM_detail] where (id_procedure = '" & DDE.GetFieldByName("id_procedure") & "')") = True Then

                Do

                    If .EOF = True Then Exit Do
                    SendDataToServer "insert into FRM_detail (id_procedure, id_analisa) values ('" & DDE.GetFieldByName("id_procedure") & "','" & DgDetail.Columns("id_analisa") & "')"
                    .MoveNext
                Loop

            End If

            .MoveLast
            DgDetail.Refresh
        End If

    End With

End Function

Private Sub Form_Load()

    With DDE
        Set .BindForm = Me
        .BindFormTAG = "ext"
        Set .ActiveConnection = CNN
        .PrepareQuery = "select * from FRM_HEADER, FRM_PROCEDURE WHERE frm_header.id_procedure = frm_procedure.id_procedure"
    End With

End Sub

Private Sub Manalisa_RowColChange(ByVal TagForm As String, _
                                  ByVal pRecordset As ADODB.Recordset)
    DDE.ChildRecordset.Fields("id_analisa") = rsanalisa.DBRecordset.Fields("id_analisa")
    DDE.ChildRecordset.Fields("nama_analisa") = rsanalisa.DBRecordset.Fields("nama_analisa")
End Sub

Private Sub Mprocedure_RowColChange(ByVal TagForm As String, _
                                    ByVal pRecordset As ADODB.Recordset)
    DDE.GetFieldByName("id_procedure") = rsprocedure.DBRecordset.Fields("id_procedure")
    Text1(0).Text = rsprocedure.DBRecordset.Fields("nama_procedure")
    Text1(2).Text = rsprocedure.DBRecordset.Fields("id_procedure")
End Sub
