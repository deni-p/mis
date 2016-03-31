VERSION 5.00
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmcontact 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contact"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8880
   Icon            =   "frmcontact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   8880
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   6
      Top             =   2385
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2430
      Left            =   0
      ScaleHeight     =   2430
      ScaleWidth      =   8880
      TabIndex        =   5
      Top             =   0
      Width           =   8880
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "PartnerID"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1815
         MaxLength       =   15
         TabIndex        =   0
         Tag             =   "Partner"
         Top             =   105
         Width           =   3195
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "CompanyName"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1815
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Partner"
         Top             =   536
         Width           =   3195
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "ContactName"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1815
         MaxLength       =   30
         TabIndex        =   2
         Tag             =   "Partner"
         Top             =   967
         Width           =   3195
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Email"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         Left            =   1815
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "Partner"
         Top             =   1398
         Width           =   3195
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "URL"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   12
         Left            =   1815
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "Partner"
         Top             =   1830
         Width           =   3195
      End
      Begin VB.Line Line1 
         Index           =   12
         X1              =   165
         X2              =   1935
         Y1              =   2145
         Y2              =   2145
      End
      Begin VB.Line Line1 
         Index           =   11
         X1              =   165
         X2              =   1935
         Y1              =   1713
         Y2              =   1713
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   165
         X2              =   2160
         Y1              =   1282
         Y2              =   1282
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   165
         X2              =   1935
         Y1              =   851
         Y2              =   851
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   165
         X2              =   1935
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   9
         Left            =   195
         TabIndex        =   11
         Top             =   1470
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Web Site"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   8
         Left            =   195
         TabIndex        =   10
         Top             =   1905
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Person"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   195
         TabIndex        =   9
         Top             =   1035
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Perusahaan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   8
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Partner ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   7
         Top             =   180
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmcontact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MyData As clsMaster


Private Sub Form_Load()
HiasFormManTell Picture2, Me
Set MyData = New clsMaster
Set MyDDE.ActiveConnection = CNN
Set MyDDE.BindForm = Me
    MyDDE.BindFormTAG = "Partner"
    MyDDE.PrepareQuery = "Select * from PartnerDB where PartnerType='Customer' Order By PartnerID"
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
Case tmbAddNew
      MyDDE.GetFieldByName("PartnerID") = MyData.PrepareIndex(tmbCustomer, 7, "", "CS")
End Select
End Sub



Private Sub PrepareQuery()
Dim TypTr As String
On Error GoTo xErr
With MyDDE
    TypTr = "CUSTOMER"
   
   .PrepareAppend = " INSERT INTO PartnerDB (PartnerID, CompanyName, ContactName, email,URL, PartnerType) " & _
                    " VALUES (N'" & ValidString(txtBox(0)) & "', N'" & ValidString(txtBox(1)) & "', N'" & ValidString(txtBox(2)) & "' ," & _
                    " N'" & ValidString(txtBox(11)) & "', N'" & ValidString(txtBox(12)) & "', N'" & TypTr & "')"
                    
   .PrepareUpdate = " UPDATE    PartnerDB Set CompanyName=N'" & ValidString(txtBox(1)) & "', ContactName = N'" & ValidString(txtBox(2)) & "'," & _
                    " Email = N'" & ValidString(txtBox(11)) & "', URL = N'" & ValidString(txtBox(12)) & "' WHERE  (PartnerID = N'" & ValidString(txtBox(0)) & "') AND (PartnerType = N'" & TypTr & "')"
                    
   .PrepareDelete = " DELETE FROM PartnerDB WHERE   (PartnerType = N'" & TypTr & "') AND (PartnerID = N'" & ValidString(txtBox(0)) & "')"
End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear
Exit Sub
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
Case tmbSave
   MyDDE.IsChildMemberReady = True
   PrepareQuery
End Select
End Sub

