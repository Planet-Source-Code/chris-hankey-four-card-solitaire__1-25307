Attribute VB_Name = "MDeclares"
Public Type RECT
    left As Long
    Top As Long
    right As Long
    bottom As Long
End Type
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptx As Long, ByVal pty As Long) As Long


Public Const HH_DISPLAY_TOPIC = &H0
Public Const HH_HELP_FINDER = &H0              ' WinHelp equivalent
Public Const HH_DISPLAY_TOC = &H1              ' not currently implemented
Public Const HH_DISPLAY_INDEX = &H2            ' not currently implemented
Public Const HH_DISPLAY_SEARCH = &H3           ' not currently implemented
Public Const HH_SET_WIN_TYPE = &H4
Public Const HH_GET_WIN_TYPE = &H5
Public Const HH_GET_WIN_HANDLE = &H6
Public Const HH_GET_INFO_TYPES = &H7           ' not currently implemented
Public Const HH_SET_INFO_TYPES = &H8           ' not currently implemented
Public Const HH_SYNC = &H9
Public Const HH_ADD_NAV_UI = &HA               ' not currently implemented
Public Const HH_ADD_BUTTON = &HB               ' not currently implemented
Public Const HH_GETBROWSER_APP = &HC           ' not currently implemented
Public Const HH_KEYWORD_LOOKUP = &HD
Public Const HH_DISPLAY_TEXT_POPUP = &HE      ' display string resource id or text in a
Public Const HH_HELP_CONTEXT = &HF            ' display mapped numeric value in dwData
Public Const HH_TP_HELP_WM_HELP = &H11         ' text popup help
Public Const HH_CLOSE_ALL = &H12               ' close all windows opened directly or
Public Const HH_ALINK_LOOKUP = &H13           ' ALink version of HH_KEYWORD_LOOKUP


'HtmlHelp api call
Public Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal _
hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, _
dwData As Any) As Long


