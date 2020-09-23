Attribute VB_Name = "BalloonAPI"
'--------------------------------------
' ________  Copyright EAguirre (c)1999
'(        ) eaguirre@comtrade.com.mx
'(  ______)
' \/
' BalloonToolTip
'--------------------------------------
Option Explicit

'Type Declarations
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type

Public Type Msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

'Window Messages
Public Const WM_MOUSEMOVE = &H200
Public Const WM_SETCURSOR = &H20
Public Const WM_HSCROLL = &H114
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_VSCROLL = &H115
'Drawing Text
Public Const DT_CALCRECT = &H400
Public Const DT_CENTER = &H1
Public Const DT_LEFT = &H0
Public Const DT_RIGHT = &H2
Public Const DT_WORDBREAK = &H10
'Region
Public Const RGN_OR = 2

'Functions Declares
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, _
                        ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, _
                        ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, _
                        ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, _
                        ByVal Y3 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, _
                        ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, _
                        ByVal nCombineMode As Long) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, _
                        ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, _
                        ByVal bRedraw As Boolean) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, _
                        ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, _
                        ByVal wFormat As Long) As Long
Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, _
                        ByVal hwnd As Long, ByVal wMsgFilterMin As Long, _
                        ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

