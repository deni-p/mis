VERSION 5.00
Object = "{43E6F32B-2B03-46D3-9276-69426FE6D51B}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmPMIXING 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DATA MIXING CHIPS MENJADI PRE LOT CHIPS"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13275
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   13275
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1500
      Index           =   0
      Left            =   120
      ScaleHeight     =   1470
      ScaleWidth      =   9585
      TabIndex        =   1
      Top             =   240
      Width           =   9615
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   1365
         Left            =   60
         ScaleHeight     =   1335
         ScaleWidth      =   9435
         TabIndex        =   2
         Top             =   60
         Width           =   9465
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            DataField       =   "m_sesudah_mixing"
            Height          =   375
            Index           =   6
            Left            =   8160
            TabIndex        =   18
            Tag             =   "mixing"
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            DataField       =   "m_sebelum_mixing"
            Height          =   375
            Index           =   5
            Left            =   8160
            TabIndex        =   17
            Tag             =   "mixing"
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            DataField       =   "k_sesudah_mixing"
            Height          =   375
            Index           =   4
            Left            =   5760
            TabIndex        =   16
            Tag             =   "mixing"
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            DataField       =   "k_sebelum_mixing"
            Height          =   375
            Index           =   3
            Left            =   5760
            TabIndex        =   15
            Tag             =   "mixing"
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            DataField       =   "total_waktu"
            Height          =   375
            Index           =   2
            Left            =   1920
            TabIndex        =   14
            Tag             =   "mixing"
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            DataField       =   "mixing_chips"
            Height          =   375
            Index           =   1
            Left            =   1920
            TabIndex        =   13
            Tag             =   "mixing"
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            DataField       =   "no_ekstrasi"
            Height          =   375
            Index           =   0
            Left            =   1920
            TabIndex        =   12
            Tag             =   "mixing"
            Top             =   120
            Width           =   1095
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   8640
            X2              =   6960
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   8640
            X2              =   6960
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   5880
            X2              =   4200
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   5880
            X2              =   4200
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   1920
            X2              =   240
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   1920
            X2              =   240
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   1920
            X2              =   240
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Sesudah Mixing"
            Height          =   255
            Index           =   8
            Left            =   6960
            TabIndex        =   11
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Sebelum Mixing"
            Height          =   255
            Index           =   7
            Left            =   6960
            TabIndex        =   10
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Moisture Chips"
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
            Index           =   6
            Left            =   6960
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Kuantitas Chips"
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
            Index           =   5
            Left            =   4200
            TabIndex        =   8
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Sesudah Mixing"
            Height          =   255
            Index           =   4
            Left            =   4200
            TabIndex        =   7
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Sebelum Mixing"
            Height          =   255
            Index           =   3
            Left            =   4200
            TabIndex        =   6
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Total waktu proses"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   5
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Mixing Chips"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   4
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "No Ekstrasi"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Width           =   855
         End
      End
   End
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   7155
      Width           =   13275
      _ExtentX        =   23416
      _ExtentY        =   1005
      BindFormTAG     =   "mixing"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmPMIXING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
   Case tmbDelete:
   DDE.PrepareDelete = "delete from mixing where no_ekstrasi = '" & DDE.GetFieldByName("no_ekstrasi") & "'"
   'Debug.Print "delete from mixing where no_ekstrasi = '" & DDE.GetFieldByName("no_ekstrasi") & "'"
End Select
End Sub

Private Sub DDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
DDE.PrepareDelete = "delete from mixing where no_ekstrasi = '" & DDE.GetFieldByName("no_ekstrasi") & "'"
End Sub

Private Sub DDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
Case tmbSave:
   DDE.IsChildMemberReady = True
   DDE.PrepareAppend = "insert into mixing (no_ekstrasi, mixing_chips, total_waktu,k_sebelum_mixing,k_sesudah_mixing, m_sebelum_mixing, m_sesudah_mixing) " & _
                       " Values ('" & DDE.GetFieldByName("no_ekstrasi") & "','" & DDE.GetFieldByName("mixing_chips") & "','" & DDE.GetFieldByName("total_waktu") & "', " & _
                       " '" & DDE.GetFieldByName("k_sebelum_mixing") & "', '" & DDE.GetFieldByName("k_sesudah_mixing") & "', " & _
                       " '" & DDE.GetFieldByName("m_sebelum_mixing") & "', '" & DDE.GetFieldByName("m_sesudah_mixing") & "')"

  Case tmbDelete:
   DDE.PrepareDelete = "delete from mixing where no_ekstrasi = '" & DDE.GetFieldByName("no_ekstrasi") & "'"

End Select
End Sub

Private Sub Form_Load()
With DDE
Set .BindForm = Me
    .BindFormTAG = "mixing"
Set .ActiveConnection = CNN
    .PrepareQuery = "select * from MIXING"
End With
End Sub

