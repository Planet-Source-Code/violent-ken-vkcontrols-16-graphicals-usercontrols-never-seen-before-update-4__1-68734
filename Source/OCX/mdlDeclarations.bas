Attribute VB_Name = "mdlDeclarations"
' =======================================================
'
' vkUserControlsXP
' Coded by violent_ken (Alain Descotes)
'
' =======================================================
'
' Some graphical UserControls for your VB application.
'
' Copyright © 2006-2007 by Alain Descotes.
'
' vkUserControlsXP is free software; you can redistribute it and/or
' modify it under the terms of the GNU Lesser General Public
' License as published by the Free Software Foundation; either
' version 2.1 of the License, or (at your option) any later version.
'
' vkUserControlsXP is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
' Lesser General Public License for more details.
'
' You should have received a copy of the GNU Lesser General Public
' License along with this library; if not, write to the Free Software
' Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
'
' =======================================================


Option Explicit



'=======================================================
'CONSTANTES
'=======================================================
Public Const MK_LBUTTON                     As Long = &H1&
Public Const MK_RBUTTON                     As Long = &H2&
Public Const MK_SHIFT                       As Long = &H4&
Public Const MK_CONTROL                     As Long = &H8&
Public Const MK_MBUTTON                     As Long = &H10&
Public Const WM_KEYDOWN                     As Long = &H100
Public Const WM_KEYFIRST                    As Long = &H100
Public Const WM_KEYLAST                     As Long = &H108
Public Const WM_KEYUP                       As Long = &H101
Public Const WM_LBUTTONDBLCLK               As Long = &H203
Public Const WM_LBUTTONDOWN                 As Long = &H201
Public Const WM_LBUTTONUP                   As Long = &H202
Public Const WM_MBUTTONDBLCLK               As Long = &H209
Public Const WM_MBUTTONDOWN                 As Long = &H207
Public Const WM_MBUTTONUP                   As Long = &H208
Public Const WM_MOUSEFIRST                  As Long = &H200
Public Const WM_MOUSELAST                   As Long = &H209
Public Const WM_MOUSEHOVER                  As Long = &H2A1
Public Const WM_MOUSELEAVE                  As Long = &H2A3
Public Const WM_MOUSEMOVE                   As Long = &H200
Public Const WM_RBUTTONDBLCLK               As Long = &H206
Public Const WM_RBUTTONDOWN                 As Long = &H204
Public Const WM_RBUTTONUP                   As Long = &H205
Public Const WM_MOUSEWHEEL                  As Long = &H20A
Public Const WM_PAINT                       As Long = &HF
Public Const GWL_WNDPROC                    As Long = -4&
Public Const TME_LEAVE                      As Long = &H2&
Public Const TME_HOVER                      As Long = &H1&
Public Const DT_CENTER                      As Long = &H1&
Public Const DT_LEFT                        As Long = &H0&
Public Const DT_RIGHT                       As Long = &H2&
Public Const DI_MASK                        As Long = &H1
Public Const DI_IMAGE                       As Long = &H2
Public Const DI_NORMAL                      As Long = DI_MASK Or DI_IMAGE
Public Const SRCCOPY                        As Long = 13369376
Public Const TIME_ONESHOT                   As Long = 0
Public Const TIME_PERIODIC                  As Long = 1
Public Const TIME_CALLBACK_EVENT_PULSE      As Long = &H20
Public Const TIME_CALLBACK_EVENT_SET        As Long = &H10
Public Const TIME_CALLBACK_FUNCTION         As Long = &H0
Public Const NIF_ICON                       As Long = &H2
Public Const NIF_MESSAGE                    As Long = &H1
Public Const HTCAPTION                      As Long = 2
Public Const NIM_DELETE                     As Long = &H2
Public Const NIF_TIP                        As Long = &H4
Public Const NIM_ADD                        As Long = &H0
Public Const NIM_MODIFY                     As Long = &H1
Public Const NIF_INFO                       As Long = &H10
Private Const GRADIENT_FILL_RECT_H          As Long = &H0
Private Const GRADIENT_FILL_RECT_V          As Long = &H1



'=======================================================
'APIs
'=======================================================
Public Declare Sub PathStripPath Lib "shlwapi.dll" Alias "PathStripPathA" (ByVal pszPath As String)
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENTTYPE) As Long
Public Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function GetTabbedTextExtent Lib "user32" Alias "GetTabbedTextExtentA" (ByVal hDC As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long) As Long
Public Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function FrameRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function timeKillEvent Lib "winmm.dll" (ByVal uID As Long) As Long
Public Declare Function timeSetEvent Lib "winmm.dll" (ByVal uDelay As Long, ByVal uResolution As Long, ByVal lpFunction As Long, ByVal dwUser As Long, ByVal uFlags As Long) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "Shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function GetActiveWindow Lib "user32.dll" () As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (ByRef lpLogFont As LOGFONT) As Long
Public Declare Function GetObjectA Lib "gdi32" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Boolean, ByVal fdwUnderline As Boolean, ByVal fdwStrikeOut As Boolean, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function DrawTextW Lib "user32.dll" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long ' Modified


Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long



'=======================================================
'TYPES
'=======================================================
Public Type TRACKMOUSEEVENTTYPE
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type
Public Type POINTAPI
    x As Long
    y As Long
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type RGB_COLOR
    R As Long
    G As Long
    B As Long
End Type
Public Type RGB_COLOR_INT
    R As Integer
    G As Integer
    B As Integer
End Type
Public Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
End Type
Public Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 128
   dwState As Long
   dwStateMask As Long
   szInfo As String * 256
   uTimeout As Long
   szInfoTitle As String * 64
   dwInfoFlags As Long
End Type
Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To 32) As Byte
End Type
Private Type TRIVERTEX
    x As Long
    y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type
Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type


'=======================================================
'FUNCTIONS PUBLIQUES
'=======================================================
'=======================================================
'conversion de nombres
'=======================================================
Public Function LoWord(DWord As Long) As Long
    If DWord And &H8000& Then ' &H8000& = &H00008000
        LoWord = DWord Or &HFFFF0000
    Else
        LoWord = DWord And &HFFFF&
    End If
End Function
Public Function HiWord(DWord As Long) As Long
    HiWord = (DWord And &HFFFF0000) \ &H10000
End Function
'=======================================================
'convertit une couleur en long vers RGB
'=======================================================
Public Sub LongToRGB(ByVal Color As Long, ByRef R As Long, ByRef G As Long, ByRef B As Long)
    R = Color And &HFF&
    G = (Color And &HFF00&) \ &H100&
    B = Color \ &H10000
End Sub
'=======================================================
'convertit une couleur en long vers RGB (Integer)
'=======================================================
Public Function LongToRGBint(ByVal Color As Long) As RGB_COLOR_INT
    LongToRGBint.R = GetCompatibleColor(Color And &HFF&)
    LongToRGBint.G = GetCompatibleColor((Color And &HFF00&) \ &H100&)
    LongToRGBint.B = GetCompatibleColor(Color \ &H10000)
End Function
'=======================================================
'convertit une couleur 0-255 vers un format compatible pour l'API GradientFillRect
'=======================================================
Private Function GetCompatibleColor(Color As Long) As Integer
    If Color > 127 Then
        GetCompatibleColor = 256 * (Color - 256)
    Else
        GetCompatibleColor = 256 * Color
    End If
End Function

'=======================================================
'création de gradients de couleur
'=======================================================
Public Sub FillGradient(DC As Long, LeftColor As Long, RightColor As _
    Long, ByVal Width As Long, ByVal Height As Long, ByVal _
    tType As GradientConstants, Optional ByVal Dep As Long = 0)
    
Dim VERT(1) As TRIVERTEX
Dim gRect As GRADIENT_RECT
Dim LC As RGB_COLOR_INT
Dim RC As RGB_COLOR_INT
Dim Hconst As Long
    
    'on convertit les Long en RGBint
    LC = LongToRGBint(LeftColor)
    RC = LongToRGBint(RightColor)
    
    'couleur de gauche que l'on définit
    With VERT(0)
        .x = 0
        .y = Dep
        .Red = LC.R
        .Green = LC.G
        .Blue = LC.B
        .Alpha = 0&
    End With

    'to blue
    With VERT(1)
        .x = Width
        .y = Height
        .Red = RC.R
        .Green = RC.G
        .Blue = RC.B
        .Alpha = 0&
    End With
    
    'on définit le type de gradient à appliquer
    If tType = Vertical Then
        Hconst = GRADIENT_FILL_RECT_H
    Else
        Hconst = GRADIENT_FILL_RECT_V
    End If
    
    With gRect
        .LowerRight = 0
        .UpperLeft = 1
    End With
    
    'maintenant on trace le gradient !
    Call GradientFillRect(DC, VERT(0), 2, gRect, 1, Hconst)

End Sub
