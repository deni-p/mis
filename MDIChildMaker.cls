VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MDIChildMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
' www.cyberbiz.com - mje@cyberbiz.com - PH: 773-338-3755
'
'LICENSE:
'--------
' You have a royalty-free right to use, modify, reproduce and distribute all
' MDIChildMaker files.
' By using any portion of MDIChildMaker you agree that CyberBiz, as well as all
' developers and distributors of MDIChildMaker have no warranty, obligations
' or liability for any MDIChildMaker related file (code, executable, etc.).
'-------------------------------------------------------------------------------

'Check the README1st.txt file in related documents for information on this application.

Option Explicit

Implements ISubclass

Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000
Private Const WS_THICKFRAME = &H40000
Private Const HWND_TOPMOST = -1
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_FORCE_REDRAW = SWP_FRAMECHANGED Or _
SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Declare Function GetWindowLong Lib _
"user32" Alias "GetWindowLongA" _
(ByVal hwnd As Long, _
ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong _
Lib "user32" Alias "SetWindowLongA" _
(ByVal hwnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long

Private Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, _
ByVal hWndInsertAfter As Long, _
ByVal X As Long, _
ByVal Y As Long, _
ByVal cx As Long, _
ByVal cy As Long, _
ByVal wFlags As Long) As Long


'-----------------------------------------------------------------
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Const WM_NCACTIVATE = &H86
Private Const WM_ACTIVATEAPP = &H1C
Private Const WM_ACTIVATE = &H6
Private Const WM_MOUSEACTIVATE = &H21
Private Const WM_SYSCOMMAND = &H112

'-----------------------------------------------------------------

Private m_oParent As Form
Private m_oChild As Form
Private m_PrevParent As Long

Public Property Set ParentForm(oForm As Form)
    Set m_oParent = oForm
End Property

Public Property Get ParentForm() As Form
    Set ParentForm = m_oParent
End Property

Public Property Set ChildForm(oForm As Form)
    Set m_oChild = oForm
End Property

Public Property Get ChildForm() As Form
    Set ChildForm = m_oChild
End Property

Public Sub MakeMDIChild()
    Dim L As Long
    
    Me.Attach
    
    m_oChild.Top = 0
    m_oChild.Left = 0
    m_oChild.Width = m_oParent.ScaleWidth
    m_oChild.Height = m_oParent.ScaleHeight
    m_oChild.Show
    
    SetParent m_oChild.hwnd, m_oParent.hwnd
    
    L = GetWindowLong(m_oChild.hwnd, GWL_STYLE)
    L = L And Not (WS_CAPTION Or WS_THICKFRAME)
    Call SetWindowLong(m_oChild.hwnd, GWL_STYLE, L)
    Call SetWindowPos(m_oChild.hwnd, _
    HWND_TOPMOST, 0, 0, 0, 0, SWP_FORCE_REDRAW)
End Sub

Public Sub Resize()
    On Error Resume Next
    
    Call MoveWindow(m_oChild.hwnd, 0, 0, _
        m_oParent.ScaleWidth, m_oParent.ScaleHeight, True)

    'An error may occur here, since we will make this method optional...
    '...that's why we trap for errors with the above Error handler
    m_oChild.ResizeCtls
End Sub

Private Sub Class_Terminate()
    Me.Detach
    
    Unload m_oParent
    Unload m_oChild
    
    Set m_oParent = Nothing
    Set m_oChild = Nothing
End Sub


Public Sub Attach()
    AttachMessage Me, m_oChild.hwnd, WM_SYSCOMMAND
    AttachMessage Me, m_oChild.hwnd, WM_MOUSEACTIVATE
    AttachMessage Me, m_oChild.hwnd, WM_NCACTIVATE
    AttachMessage Me, m_oChild.hwnd, WM_ACTIVATE
End Sub


Public Sub Detach()
    DetachMessage Me, m_oChild.hwnd, WM_SYSCOMMAND
    DetachMessage Me, m_oChild.hwnd, WM_MOUSEACTIVATE
    DetachMessage Me, m_oChild.hwnd, WM_NCACTIVATE
    DetachMessage Me, m_oChild.hwnd, WM_ACTIVATE
    
    Call SetParent(m_oChild.hwnd, GetDesktopWindow())
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)

End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
   ' Determine the response to the current message:
   Select Case CurrentMessage
   Case WM_SYSCOMMAND
        ISubclass_MsgResponse = EMsgResponse.emrConsume
   Case WM_NCACTIVATE
      ISubclass_MsgResponse = emrConsume
   Case Else
      ISubclass_MsgResponse = emrPostProcess
   End Select
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Static s_iRefCount As Long
    Select Case iMsg
        'Don't let the cgild window accept commands, ex. like alt+F4"close"
        Case WM_SYSCOMMAND
           'Ignore the messages...
           'TODO - This doesn't work...How do we pass the messages to the parent form???
           'SendMessage m_oParent.hwnd, iMsg, wParam, lParam
        Case WM_MOUSEACTIVATE
           BringWindowToTop m_oParent.hwnd
           ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
        Case WM_NCACTIVATE
           ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
        Case WM_ACTIVATE
            'Make sure the main form stays active when the MDIChild is activated/deactivated
           If (wParam = 0) Then ' app being deactivated
                SendMessage GetParent(GetParent(m_oParent.hwnd)), WM_NCACTIVATE, 0, 0
           Else
                SendMessage GetParent(GetParent(m_oParent.hwnd)), WM_NCACTIVATE, 1, 0
           End If
    End Select
End Function


