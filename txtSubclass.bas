Attribute VB_Name = "txtSubclass"
'Subclass A TextBox
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Const GWL_EXSTYLE = (-20)

Public Const GWL_WNDPROC = -4
    Public Const WM_RBUTTONUP = &H205
    Private Const WM_COPY As Long = &H301
    Private Const WM_PASTE As Long = &H302
    Global lpPrevWndProc As Long
    Global gHW As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Sub Hook()
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Public Sub UnHook()
    Dim lngReturnValue As Long
    lngReturnValue = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub
Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_RBUTTONUP
            MsgBox "SubClassing Text Box, Whether is it right or wrong process let me know from you. Comment On this at debughosh@vsnl.net", vbCritical
        Case WM_COPY
            MsgBox "SubClassing Text Box, Whether is it right or wrong process let me know from you. Comment On this at debughosh@vsnl.net", vbCritical
        Case WM_PASTE
            MsgBox "SubClassing Text Box, Whether is it right or wrong process let me know from you. Comment On this at debughosh@vsnl.net", vbCritical
        Case Else
            WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
        End Select
End Function

