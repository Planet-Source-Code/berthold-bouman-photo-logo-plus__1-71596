Attribute VB_Name = "modAPI"
Option Explicit

'change progbar colors
Public Const WM_USER = &H400
Public Const CCM_FIRST = &H2000&                'common control shared messages
Public Const CCM_SETBKCOLOR = (CCM_FIRST + 1)   'lParam = bkColor
Public Const PBM_SETBKCOLOR = CCM_SETBKCOLOR    'lParam = bkColor (IE3 & later)
Public Const PBM_SETBARCOLOR = (WM_USER + 9)    'lParam = barcolor (IE4 & later)

Public Declare Function BitBlt Lib "gdi32" _
                (ByVal hDestDC As Long, _
                 ByVal x As Long, _
                 ByVal y As Long, _
                 ByVal nWidth As Long, _
                 ByVal nHeight As Long, _
                 ByVal hSrcDC As Long, _
                 ByVal xSrc As Long, _
                 ByVal ySrc As Long, _
                 ByVal dwRop As Long) As Long

Public Declare Function TransparentBlt Lib "msimg32.dll" _
                (ByVal hDC As Long, _
                 ByVal x As Long, _
                 ByVal y As Long, _
                 ByVal nWidth As Long, _
                 ByVal nHeight As Long, _
                 ByVal hSrcDC As Long, _
                 ByVal xSrc As Long, _
                 ByVal ySrc As Long, _
                 ByVal nSrcWidth As Long, _
                 ByVal nSrcHeight As Long, _
                 ByVal crTransparent As Long) As Boolean

Public Declare Function StretchBlt Lib "gdi32" _
                (ByVal hDC As Long, _
                 ByVal x As Long, _
                 ByVal y As Long, _
                 ByVal nWidth As Long, _
                 ByVal nHeight As Long, _
                 ByVal hSrcDC As Long, _
                 ByVal xSrc As Long, _
                 ByVal ySrc As Long, _
                 ByVal nSrcWidth As Long, _
                 ByVal nSrcHeight As Long, _
                 ByVal dwRop As Long) As Long
                 
Public Const ScrCopy = &HCC0020

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'supports make picture transparent routine
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, _
                        ByVal y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, _
                        ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, _
                        ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, _
                        ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, _
                        ByVal nCombineMode As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Const RGN_OR = 2

Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Type POINTAPI
    x As Long
    y As Long
End Type

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'move form without caption
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                        ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'open browse for folder dialog
Public Type BrowseInfo
     hwndOwner As Long
     pIDLRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hmem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'alphablending
Public Declare Function AlphaBlend Lib "msimg32.dll" ( _
            ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
            ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
            ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, _
            ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
            Destination As Any, Source As Any, ByVal Length As Long)

'type structure
Public Type typeBlendProperties
    tBlendOp As Byte
    tBlendOptions As Byte
    tBlendAmount As Byte
    tAlphaType As Byte
End Type




