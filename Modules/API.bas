Attribute VB_Name = "API"
Option Explicit

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Public Const API_TRUE As Long = 1&
Public Const API_FALSE As Long = 0&

'***[Windows Messages]**************************************************************
Public Const WM_NULL             As Long = &H0
Public Const WM_CREATE           As Long = &H1
Public Const WM_DESTROY          As Long = &H2
Public Const WM_MOVE             As Long = &H3
Public Const WM_SIZE             As Long = &H5
Public Const WM_ACTIVATE         As Long = &H6
Public Const WM_NCMOUSEMOVE      As Long = &HA0
Public Const WM_NCLBUTTONDOWN    As Long = &HA1
Public Const WM_NCLBUTTONUP      As Long = &HA2
Public Const WM_NCLBUTTONDBLCLK  As Long = &HA3
Public Const WM_NCRBUTTONDOWN    As Long = &HA4
Public Const WM_NCMBUTTONDOWN    As Long = &HA7
Public Const WM_NCMBUTTONUP      As Long = &HA8
Public Const WM_NCMBUTTONDBLCLK  As Long = &HA9
Public Const WM_NCHITTEST        As Long = &H84
Public Const WM_SYSCOMMAND       As Long = &H112
Public Const WM_MOUSEMOVE        As Long = &H200
Public Const WM_LBUTTONDOWN      As Long = &H201
Public Const WM_LBUTTONUP        As Long = &H202
Public Const WM_LBUTTONDBLCLK    As Long = &H203
Public Const WM_RBUTTONDOWN      As Long = &H204
Public Const WM_RBUTTONUP        As Long = &H205
Public Const WM_RBUTTONDBLCLK    As Long = &H206
Public Const WM_MBUTTONDOWN      As Long = &H207
Public Const WM_MBUTTONUP        As Long = &H208
Public Const WM_MBUTTONDBLCLK    As Long = &H209
Public Const WM_MOUSEWHEEL       As Long = &H20A
Public Const WM_PAINT            As Long = &HF
Public Const WM_USER             As Long = &H400
Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115

'Message parameters
Public Const HTBORDER = 18
Public Const HTBOTTOM = 15
Public Const HTBOTTOMLEFT = 16
Public Const HTBOTTOMRIGHT = 17
Public Const HTCAPTION = 2
Public Const HTCLIENT = 1
Public Const HTERROR = (-2)
Public Const HTGROWBOX = 4
Public Const HTHSCROLL = 6
Public Const HTLEFT = 10
Public Const HTMAXBUTTON = 9
Public Const HTMENU = 5
Public Const HTMINBUTTON = 8
Public Const HTNOWHERE = 0
Public Const HTREDUCE = HTMINBUTTON
Public Const HTRIGHT = 11
Public Const HTSIZE = HTGROWBOX
Public Const HTSIZEFIRST = HTLEFT
Public Const HTSIZELAST = HTBOTTOMRIGHT
Public Const HTSYSMENU = 3
Public Const HTTOP = 12
Public Const HTTOPLEFT = 13
Public Const HTTOPRIGHT = 14
Public Const HTTRANSPARENT = (-1)
Public Const HTVSCROLL = 7
Public Const HTZOOM = HTMAXBUTTON

'***[Windows]*******************************************************************
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long

'***[Focus Rectangle]**************************************************************
Public Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long

'***[Mouse Pointer]**************************************************************
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'***[ScrollBars]**************************************************************
Public Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

'ScrollBar Constants
Public Const SB_HORZ = 0
Public Const SB_VERT = 1
Public Const SB_CTL = 2

Public Const SIF_RANGE = &H1
Public Const SIF_PAGE = &H2
Public Const SIF_POS = &H4
Public Const SIF_DISABLENOSCROLL = &H8
Public Const SIF_TRACKPOS = &H10
Public Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

'ScrollBar Commands
Public Const SB_LINEUP As Long = 0&
Public Const SB_LINELEFT As Long = 0&
Public Const SB_LINEDOWN As Long = 1&
Public Const SB_LINERIGHT As Long = 1&
Public Const SB_PAGEUP As Long = 2&
Public Const SB_PAGELEFT As Long = 2&
Public Const SB_PAGEDOWN As Long = 3&
Public Const SB_PAGERIGHT As Long = 3&
Public Const SB_THUMBPOSITION As Long = 4&
Public Const SB_THUMBTRACK As Long = 5&
Public Const SB_TOP As Long = 6&
Public Const SB_LEFT As Long = 6&
Public Const SB_BOTTOM As Long = 7&
Public Const SB_RIGHT As Long = 7&
Public Const SB_ENDSCROLL As Long = 8&


Public Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
Public Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Public Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long

'***[Text]********************************************************************
Public Const DT_LEFT = &H0
Public Const DT_CENTER = &H1
Public Const DT_RIGHT = &H2
Public Const DT_WORDBREAK = &H10
Public Const DT_PATH_ELLIPSIS = &H4000
Public Const DT_END_ELLIPSIS = &H8000

Public Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type

Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
'Public Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Public Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, ByVal lpDrawTextParams As Any) As Long

'***[16-bit Integer extraction]********************************************************************
Public Function LOWORD(ByVal Value As Long) As Integer
    'Returns the low 16-bit integer from a 32-bit long integer
    CopyMemory LOWORD, Value, 2&
End Function

Public Function HIWORD(ByVal Value As Long) As Integer
    'Returns the high 16-bit integer from a 32-bit long integer
    CopyMemory HIWORD, ByVal VarPtr(Value) + 2, 2&
End Function

