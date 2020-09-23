Attribute VB_Name = "modWndProc"
Option Explicit
'

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_WNDPROC = -4

Private mvarPrevWndProc As Long
Private mvarControl As SuperSubClasser

Public Sub Hook(ByVal lHwnd As Long, lObject As SuperSubClasser)
    mvarPrevWndProc = SetWindowLong(lHwnd, GWL_WNDPROC, AddressOf WindowProc)
    Set mvarControl = lObject
End Sub

Public Sub UnHook(ByVal lHwnd As Long)
    Dim lngReturnValue As Long
    lngReturnValue = SetWindowLong(lHwnd, GWL_WNDPROC, mvarPrevWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    mvarControl.GotMessage uMsg, wParam, lParam
    WindowProc = CallWindowProc(mvarPrevWndProc, hw, uMsg, wParam, lParam)
End Function


