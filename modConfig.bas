Option Explicit

'====================================================
' Youzhi PPT Kit
' 公共参数
'====================================================

Public Const YZK_VERSION As String = "2.0.0"

'=========================
' 字体
'=========================

Public Const FONT_CN As String = "微软雅黑"
Public Const FONT_EN As String = "Times New Roman"

Public Const FONT_STANDARD As Single = 18
Public Const FONT_EMPHASIS As Single = 24
Public Const FONT_LIST As Single = 18
Public Const FONT_CONCLUSION As Single = 24


'=========================
' 线宽
'=========================

Public Const LINE_PICTURE As Single = 0.75
Public Const LINE_BOX As Single = 4
Public Const LINE_GRAY_DASH As Single = 1
Public Const LINE_ARROW As Single = 4


'=========================
' 颜色
'=========================

Public Function YZK_Black() As Long
    YZK_Black = RGB(0, 0, 0)
End Function

Public Function YZK_White() As Long
    YZK_White = RGB(255, 255, 255)
End Function

' #C00000
Public Function YZK_Red() As Long
    YZK_Red = RGB(192, 0, 0)
End Function

' #2E54A1
Public Function YZK_Blue() As Long
    YZK_Blue = RGB(46, 84, 161)
End Function

' #E6E6E6
Public Function YZK_LightGray() As Long
    YZK_LightGray = RGB(230, 230, 230)
End Function

' #D9D9D9
Public Function YZK_Gray() As Long
    YZK_Gray = RGB(217, 217, 217)
End Function
Public Function YZK_ColorBlack() As Long
    YZK_ColorBlack = RGB(0, 0, 0)
End Function

Public Function YZK_ColorRed() As Long
    YZK_ColorRed = RGB(192, 0, 0)
End Function

Public Function YZK_ColorBlue() As Long
    YZK_ColorBlue = RGB(46, 84, 161)
End Function

Public Function YZK_ColorLightGray() As Long
    YZK_ColorLightGray = RGB(230, 230, 230)
End Function

Public Function YZK_ColorGray() As Long
    YZK_ColorGray = RGB(217, 217, 217)
End Function
