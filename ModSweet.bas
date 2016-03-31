Attribute VB_Name = "ModSweet"
'Option Explicit
'Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
''Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'
'Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
'
'Private Const MF_BYPOSITION = &H400&
'
''Private ReadyToClose As Boolean
'Private Sub RemoveMenus(frm As Form, _
'    remove_restore As Boolean, _
'    remove_move As Boolean, _
'    remove_size As Boolean, _
'    remove_minimize As Boolean, _
'    remove_maximize As Boolean, _
'    remove_seperator As Boolean, _
'    remove_close As Boolean)
'Dim hMenu As Long
'
'    ' Get the form's system menu handle.
'    hMenu = GetSystemMenu(frm.hwnd, False)
'
'    If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION
'    If remove_seperator Then DeleteMenu hMenu, 5, MF_BYPOSITION
'    If remove_maximize Then DeleteMenu hMenu, 4, MF_BYPOSITION
'    If remove_minimize Then DeleteMenu hMenu, 3, MF_BYPOSITION
'    If remove_size Then DeleteMenu hMenu, 2, MF_BYPOSITION
'    If remove_move Then DeleteMenu hMenu, 1, MF_BYPOSITION
'    If remove_restore Then DeleteMenu hMenu, 0, MF_BYPOSITION
'End Sub
'
'Public Sub SweetForm(ByVal FormName As Form, _
'                      Optional ByVal ColorPanelUp As OLE_COLOR = &H80000002, _
'                      Optional ByVal ColorPanelBottom As OLE_COLOR = &H80000001)
'Dim Rc As RECT
'Dim Obj As Object
'Dim MinHeight As Long
'For Each Obj In FormName.Controls
'    If TypeOf Obj Is PictureBox Then
'       MinHeight = Obj.Height
'    End If
'Next Obj
'FormName.Cls
'FormName.ScaleMode = vbPixels
'FormName.BackColor = vbWhite
'With Rc
'    .Left = 1
'    .Top = 1
'    .Bottom = .Top + 45
'    .Right = FormName.ScaleWidth - 2
'End With
''GetClientRect FormName.hwnd, rc
'Draw3dRect FormName.hDC, Rc, ColorPanelBottom, ColorPanelBottom
''With rc
''    .Left = 1
''    .Top = 1
''    .Bottom = .Top + 45
''    .Right = FormName.ScaleWidth - 2
''End With
''FillSolidRect FormName.hDC, rc.Left, rc.Top, rc.Right, rc.Bottom, ColorPanelUp
''With rc
''    .Left = 1
''    .Top = 31
''    .Bottom = .Top + 15
''    .Right = FormName.ScaleWidth - 2
''End With
''FillSolidRect FormName.hDC, rc.Left, rc.Top, rc.Right, rc.Bottom - 30, ColorPanelBottom ' &HD97E60
''With rc
''    .Left = 2
''    .Top = 48
''    .Bottom = FormName.ScaleHeight - (2 + MinHeight)
''    .Right = FormName.ScaleWidth - 2
''End With
''With rc
''    .Left = 6
''    .Top = 53
''    .Bottom = FormName.ScaleHeight - (30 + MinHeight)
''    .Right = FormName.ScaleWidth - 12
''End With
'
''FillSolidRect FormName.hDC, rc.Left, rc.Top, rc.Right, rc.Bottom - 30, &HF2E2D5
'FormName.ScaleMode = vbTwips
'FormName.Refresh
''RemoveMenus FormName, False, False, False, True, True, False, False
'End Sub
'
'
'Public Sub Draw3dRect(hDC As Long, Rc As RECT, clrTopLeft As OLE_COLOR, _
'    clrBottomRight As OLE_COLOR)
'    Dim X As Long, y As Long, cx As Long, cy As Long
'    X = Rc.Left
'    y = Rc.Top
'    cx = Rc.Right - Rc.Left
'    cy = Rc.Bottom - Rc.Top
'
'    FillSolidRect hDC, X, y, cx - 1, 1, clrTopLeft
'    FillSolidRect hDC, X, y, 1, cy - 1, clrTopLeft
'    FillSolidRect hDC, X + cx, y, -1, cy, clrBottomRight
'    FillSolidRect hDC, X, y + cy, cx, -1, clrBottomRight
'End Sub
'
'Private Sub FillSolidRect(hDC As Long, X As Long, y As Long, cx As Long, _
'    cy As Long, clr As OLE_COLOR)
'    Dim hBr As Long, Rc As RECT
'    Rc.Left = X
'    Rc.Top = y
'    Rc.Right = X + cx
'    Rc.Bottom = y + cy
'    hBr = CreateSolidBrush(TranslateColor(clr))
'    FillRect hDC, Rc, hBr
'    DeleteObject hBr
'End Sub
'
'Private Function TranslateColor(ByVal clr As OLE_COLOR, _
'                        Optional hPal As Long = 0) As Long
'    On Error Resume Next
'    If OleTranslateColor(clr, hPal, TranslateColor) Then
'        TranslateColor = -1
'    End If
'    Err.Clear
'End Function
'

'

'

'
'
