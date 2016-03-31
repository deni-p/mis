VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFindHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vTag, vPartner As String
Private aDate, aID, aPartner As String
Private xSelected As Boolean

Public Property Let Tag(vData As String)
   vTag = vData
End Property

Public Property Let PartnerAs(vData As String)
   vPartner = vData
End Property


Public Sub ShowForm()
   frmFindHistory.HistoryName = vTag
   frmFindHistory.PartnerType = vPartner
   
   frmFindHistory.Show vbModal
   xSelected = frmFindHistory.IsSelected
   
   aID = frmFindHistory.ID
   aDate = frmFindHistory.DateRange
   aPartner = frmFindHistory.Partner
End Sub

Public Function GetID() As String
   GetID = aID
End Function

Public Function GetDateRange() As String
   GetDateRange = aDate
End Function

Public Function GetPartner() As String
   GetPartner = aPartner
End Function

Public Function IsSelected() As Boolean
   IsSelected = xSelected
End Function

Public Sub SetHistory(ViewName As String, DetailFormName As Form, Optional Tag As String, Optional Partner As String)
   Dim frmHist As New frmHistory
   If Len(Tag) > 0 Then vTag = Tag
   If Len(Partner) > 0 Then vPartner = Partner
   ShowForm
   If IsSelected Then
       frmHist.PartnerName = vPartner
       frmHist.SONo = aID
       frmHist.TagName = vTag
       frmHist.RangeTanggal = aDate
       frmHist.Partner = aPartner
       frmHist.ViewName = ViewName
       frmHist.DetailForm = DetailFormName
       frmHist.SetFocus
   End If
End Sub


