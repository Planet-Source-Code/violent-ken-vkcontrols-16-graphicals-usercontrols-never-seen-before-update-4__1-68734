VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVirtualDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'VARIABLES PRIVEES
'=======================================================
Private lHeight As Long
Private lWidth As Long
Private lForeColor As Long
Private lBackColor As Long
Private tFont As StdFont
Private lhDc As Long
Private lImage As Long
Private tPicture As StdPicture
Private lRetDC As Long
Private hOldFont As Long
Private hFont As Long
Private hPen As Long
Private lPenWidth As Long
Private ttPenStyle As tPenStyle
Private lPenColor As Long
Private tBackMode As BackModeEnum
Private lBkColor As Long
Private hBrush As Long
Private tBrushStyle As BrushStyleEnum
Private tBrushHatch As BrushHatchEnum
Private lBrushColor As Long


'=======================================================
'APIS
'=======================================================
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CopyImage Lib "user32" (ByVal hImage As Long, ByVal uType As Long, ByVal OutputWidth As Long, ByVal OutputHeight As Long, ByVal fuFlags As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function InvertRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DrawTextAPI Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function GetPixelAPI Lib "gdi32" Alias "GetPixel" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DrawFocusRectAPI Lib "user32" Alias "DrawFocusRect" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function DrawEdgeAPI Lib "user32" Alias "DrawEdge" (ByVal hDC As Long, qrc As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long



'=======================================================
'CONSTANTES
'=======================================================
Private Const RASTERCAPS                    As Long = 38
Private Const SIZEPALETTE                   As Long = 104
Private Const RC_PALETTE                    As Long = &H100
Private Const HWND_DESKTOP                  As Long = 0&
Private Const LOGPIXELSX                    As Long = 88
Private Const LOGPIXELSY                    As Long = 90
Private Const GRADIENT_FILL_RECT_H          As Long = &H0
Private Const GRADIENT_FILL_RECT_V          As Long = &H1



'=======================================================
'TYPES & ENUMS
'=======================================================
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type
Private Type PALETTEENTRY
    peRed               As Byte
    peGreen             As Byte
    peBlue              As Byte
    peFlags             As Byte
End Type
Private Type LOGPALETTE
    palVersion          As Integer
    palNumEntries       As Integer
    palPalEntry(255)    As PALETTEENTRY
End Type
Private Type PicBmp
    Size                As Long
    Type                As Long
    hBmp                As Long
    hPal                As Long
    Reserved            As Long
End Type
Private Type GUID
    Data1               As Long
    Data2               As Integer
    Data3               As Integer
    Data4(7)            As Byte
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Enum eGradientConstants
    eVertical = 1
    eHorizontal = 2
End Enum
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long '0=f; 255=v
    lfItalic As Byte '0=f; 255=v
    lfUnderline As Byte '0=f; 255=v
    lfStrikeOut As Byte '0=f; 255=v
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(32) As Byte
End Type
Public Enum TextStyle
    aDT_BOTTOM = &H8&
    aDT_CENTER = &H1&
    aDT_LEFT = &H0&
    aDT_CALCRECT = &H400&
    aDT_WORDBREAK = &H10&
    aDT_VCENTER = &H4&
    aDT_TOP = &H0&
    aDT_TABSTOP = &H80&
    aDT_SINGLELINE = &H20&
    aDT_RIGHT = &H2&
    aDT_NOCLIP = &H100&
    aDT_INTERNAL = &H1000&
    aDT_EXTERNALLEADING = &H200&
    aDT_EXPANDTABS = &H40&
    aDT_CHARSTREAM = 4&
    aDT_NOPREFIX = &H800&
    aDT_EDITCONTROL = &H2000&
    aDT_PATH_ELLIPSIS = &H4000&
    aDT_END_ELLIPSIS = &H8000&
    aDT_MODIFYSTRING = &H10000
    aDT_RTLREADING = &H20000
    aDT_WORD_ELLIPSIS = &H40000
End Enum
Public Enum tPenStyle
    PS_SOLID = 0
    PS_DASH = 1                    '  -------
    PS_DOT = 2                     '  .......
    PS_DASHDOT = 3                 '  _._._._
    PS_DASHDOTDOT = 4              '  _.._.._
    PS_NULL = 5
    PS_INSIDEFRAME = 6
    PS_USERSTYLE = 7
    PS_ALTERNATE = 8
End Enum
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
Private Type RGB_COLOR_INT
    R As Integer
    G As Integer
    B As Integer
End Type
Public Enum BackModeEnum
    eTRANSPARENT = 1
    SOLID = 2
End Enum
Public Enum EdgeEnum
    BDR_RAISEDOUTER = &H1
    BDR_SUNKENOUTER = &H2
    BDR_RAISEDINNER = &H4
    BDR_SUNKENINNER = &H8
    BDR_OUTER = &H3
    BDR_INNER = &HC
    BDR_RAISED = &H5
    BDR_SUNKEN = &HA
    EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
    EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
    EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
    EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
End Enum
Public Enum EdgeFlags
    BF_LEFT = &H1
    BF_TOP = &H2
    BF_RIGHT = &H4
    BF_BOTTOM = &H8
    BF_TOPLEFT = (BF_TOP Or BF_LEFT)
    BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
    BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
    BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
    BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
    BF_DIAGONAL = &H10
    BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
    BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
    BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
    BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
    BF_MIDDLE = &H800        ' Fill in the middle
    BF_SOFT = &H1000         ' For softer buttons
    BF_ADJUST = &H2000       ' Calculate the space left over
    BF_FLAT = &H4000         ' For flat rather than 3D borders
    BF_MONO = &H8000         ' For monochrome borders
End Enum
Public Enum BrushStyleEnum
    aBS_SOLID = 0
    aBS_NULL = 1
    aBS_HOLLOW = aBS_NULL
    aBS_HATCHED = 2
    aBS_PATTERN = 3
    aBS_INDEXED = 4
    aBS_DIBPATTERN = 5
    aBS_DIBPATTERNPT = 6
    aBS_PATTERN8X8 = 7
    aBS_DIBPATTERN8X8 = 8
End Enum
Public Enum BrushHatchEnum
    aHS_HORIZONTAL = 0 ' -----
    aHS_VERTICAL = 1 ' |||||
    aHS_FDIAGONAL = 2 ' \\\\\
    aHS_BDIAGONAL = 3 ' /////
    aHS_CROSS = 4 ' +++++
    aHS_DIAGCROSS = 5 ' xxxxx
    aHS_FDIAGONAL1 = 6
    aHS_BDIAGONAL1 = 7
    aHS_SOLID = 8
    aHS_DENSE1 = 9
    aHS_DENSE2 = 10
    aHS_DENSE3 = 11
    aHS_DENSE4 = 12
    aHS_DENSE5 = 13
    aHS_DENSE6 = 14
    aHS_DENSE7 = 15
    aHS_DENSE8 = 16
    aHS_NOSHADE = 17
    aHS_HALFTONE = 18
    aHS_SOLIDCLR = 19
    aHS_DITHEREDCLR = 20
    aHS_SOLIDTEXTCLR = 21
    aHS_DITHEREDTEXTCLR = 22
    aHS_SOLIDBKCLR = 23
    aHS_DITHEREDBKCLR = 24
    aHS_API_MAX = 25
End Enum



'=======================================================
'PROPERTIES
'=======================================================
Public Property Get Width() As Long
    Width = lWidth
End Property
Public Property Get Height() As Long
    Height = lHeight
End Property
Public Property Get BackColor() As Long
    BackColor = lBackColor
End Property
Public Property Get Font() As StdFont
    Set Font = tFont
End Property
Public Property Get hDC() As Long
    hDC = lhDc
End Property
Public Property Get Picture() As StdPicture
    Set Picture = hDCToPicture(Me.hDC, 0, 0, lWidth, lHeight)
End Property

Public Property Let Width(Width As Long)
    lWidth = Width
End Property
Public Property Let Height(Height As Long)
    lHeight = Height
End Property
Public Property Let BackColor(BackColor As Long)
    lBackColor = BackColor
    'on trace un rectangle pour pouvoir remplir le fond
    Call Me.FillArea(0, 0, lWidth, lHeight, BackColor)
End Property
Public Property Let ForeColor(ForeColor As Long)
    lForeColor = ForeColor
    Call SetTextColor(Me.hDC, GetColor(ForeColor))
End Property
Public Property Let Font(Font As StdFont)
    Set tFont = Font
    Call AssociateFont(Font)
End Property
Public Property Let hDC(hDC As Long)
    lhDc = hDC
End Property
Public Property Set Picture(Picture As StdPicture)
    Set tPicture = Picture
End Property

Public Property Get PenStyle() As tPenStyle
    PenStyle = ttPenStyle
End Property
Public Property Get PenWidth() As Long
    PenWidth = lPenWidth
End Property
Public Property Get PenColor() As Long
    PenColor = lPenColor
End Property
Public Property Let PenStyle(PenStyle As tPenStyle)
    ttPenStyle = PenStyle
    Call SelectNewPen
End Property
Public Property Let PenWidth(PenWidth As Long)
    lPenWidth = PenWidth
    Call SelectNewPen
End Property
Public Property Let PenColor(PenColor As Long)
    lPenColor = PenColor
    Call SelectNewPen
End Property

Public Property Get BackMode() As BackModeEnum
    BackMode = tBackMode
End Property
Public Property Let BackMode(BackMode As BackModeEnum)
    tBackMode = BackMode
    If Me.hDC Then Call SetBkMode(Me.hDC, BackMode)
End Property
Public Property Get BkColor() As Long
    BkColor = lBkColor
End Property
Public Property Let BkColor(BkColor As Long)
    lBkColor = BkColor
    If Me.hDC Then Call SetBkColor(Me.hDC, BkColor)
End Property

Public Property Get BrushStyle() As BrushStyleEnum
    BrushStyle = tBrushStyle
End Property
Public Property Get BrushColor() As Long
    BrushColor = lBrushColor
End Property
Public Property Get BrushHatch() As BrushHatchEnum
    BrushHatch = tBrushHatch
End Property
Public Property Let BrushStyle(BrushStyle As BrushStyleEnum)
    tBrushStyle = BrushStyle
    Call SelectNewBrush
End Property
Public Property Let BrushColor(BrushColor As Long)
    lBrushColor = BrushColor
    Call SelectNewBrush
End Property
Public Property Let BrushHatch(BrushHatch As Long)
    tBrushHatch = BrushHatch
    Call SelectNewBrush
End Property


'=======================================================
'CLASS SUBS
'=======================================================
Private Sub Class_Initialize()
    
    tBackMode = TRANSPARENT
    
    'on définit un pen par défaut
    ttPenStyle = PS_SOLID
    lPenWidth = 1
    lPenColor = vbBlack
    
    'on définit un brush par défaut
    tBrushStyle = aBS_NULL
    lBrushColor = vbBlack
    tBrushHatch = False
    
    'le sélectionne
    Call SelectNewPen
    
End Sub
Private Sub Class_Terminate()
    
    'on libère tout
    Set tFont = Nothing
    Set Me.Picture = Nothing
    
    'on destroy les objets
    Call DeleteObject(hFont)
    Call DeleteObject(hOldFont)
    Call DeleteObject(lImage)
    Call DeleteObject(hPen)
    Call DeleteObject(hBrush)
    Call DeleteDC(Me.hDC)

End Sub




'=======================================================
'PUBLIC SUB/FUNCTION
'=======================================================
'=======================================================
'initialise en créant le DC compatible
'=======================================================
Public Sub CreateMyDC(Width As Long, Height As Long)
Dim hDesktop As Long

    'récupère les couleurs et les tailles
    With Me
        .Width = Width
        .Height = Height
    End With

    'récupère DC du bureau
    hDesktop = GetDC(HWND_DESKTOP)

    'création DC compatible
    Me.hDC = CreateCompatibleDC(hDesktop)

    'création image
    lImage = CreateCompatibleBitmap(hDesktop, lWidth, lHeight)
    
    'on lie l'image
    lRetDC = SelectObject(Me.hDC, lImage)

    'initialise la Picture
    Set Me.Picture = Nothing
    
    'libère le bureau
    Call ReleaseDC(HWND_DESKTOP, hDesktop)
    
    'change le BackMode
    Call SetBkMode(Me.hDC, tBackMode)
    
End Sub

'=======================================================
'copie le contenu de notre picture vers le HDC désiré
'=======================================================
Public Sub PaintToDC(hDC As Long, Left As Long, Top As Long)
    Call BitBlt(hDC, Left, Top, lWidth, lHeight, Me.hDC, 0&, 0&, vbSrcCopy)
End Sub

'=======================================================
'ajoute une image dans notre picture
'=======================================================
Public Sub DrawImageToVirtualDC(Obj As StdPicture, Left As Long, Top _
    As Long, oWidth As Long, oHeight As Long, iWidth As Long, iHeight _
    As Long)
    
Dim hDesktop As Long
Dim lBmpCopy As Long
Dim hTempDC As Long
Dim hMem As Long
    
    'récupère bureau
    hDesktop = GetDC(HWND_DESKTOP)

    'création DC compatible
    hTempDC = CreateCompatibleDC(hDesktop)

    'copie image
    lBmpCopy = CopyImage(Obj.Handle, 0&, oWidth, oHeight, 0&)

    'on lie la copie l'image
    hMem = SelectObject(hTempDC, lBmpCopy)

    'libère le bureau
    Call ReleaseDC(HWND_DESKTOP, hDesktop)

    'applique le STD
    Call BitBlt(Me.hDC, Left, Top, iWidth, iHeight, hTempDC, 0, 0, vbSrcCopy)

    'libère.
    Call SelectObject(hTempDC, hMem)
    Call DeleteDC(hTempDC)
    Call ReleaseDC(HWND_DESKTOP, hDesktop)

End Sub

'=======================================================
'print du texte dans notre picture
'=======================================================
Public Sub DrawText(strString As String, Left As Long, Top As Long, _
    Width As Long, Height As Long, Optional ByVal Font As StdFont, _
    Optional ByVal Style As TextStyle = aDT_LEFT)

Dim tF As StdFont
Dim tFont As LOGFONT
Dim tRect As RECT
    
    If Not (Font Is Nothing) Then
        'on associe une nouvelle fonte
        Call AssociateFont(Font)
    End If
    
    'définit la zone
    With tRect
        .Left = Left
        .Top = Top
        .Right = Width
        .Bottom = Height
    End With
    
    'trace le texte
    Call DrawTextAPI(Me.hDC, strString, Len(strString), tRect, Style)
    
End Sub

'=======================================================
'trace une ligne
'=======================================================
Public Sub DrawLine(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, _
    Optional ByVal FillRect As Boolean = False)

Dim T As POINTAPI
Dim hPen As Long

    If FillRect Then
        'alors il s'agit d'un tracé de zone
        Call FillArea(X1, Y1, X2, Y2, Color)
        Exit Sub
    End If
    
    'déplace vers le point actif
    Call MoveToEx(Me.hDC, X1, Y1, T)
    
    'trace la ligne
    Call LineTo(Me.hDC, X2, Y2)
    
End Sub

'=======================================================
'effectue un gradient de couleur dans une zone
'=======================================================
Public Sub DrawGradient(Color1 As Long, Color2 As Long, Left As Long, _
    Top As Long, Width As Long, Height As Long, tType As eGradientConstants)
    
Dim VERT(1) As TRIVERTEX
Dim gRect As GRADIENT_RECT
Dim LC As RGB_COLOR_INT
Dim RC As RGB_COLOR_INT
Dim Hconst As Long
    
    'on convertit les Long en RGBint
    LC = LongToRGBint(Color1)
    RC = LongToRGBint(Color2)
    
    'couleur de gauche que l'on définit
    With VERT(0)
        .x = Left
        .y = Top
        .Red = LC.R
        .Green = LC.G
        .Blue = LC.B
        .Alpha = 0&
    End With

    'to blue
    With VERT(1)
        .x = Width + Left
        .y = Height + Top
        .Red = RC.R
        .Green = RC.G
        .Blue = RC.B
        .Alpha = 0&
    End With
    
    'on définit le type de gradient à appliquer
    If tType = eVertical Then
        Hconst = GRADIENT_FILL_RECT_H
    Else
        Hconst = GRADIENT_FILL_RECT_V
    End If
    
    With gRect
        .LowerRight = 0
        .UpperLeft = 1
    End With
    
    'maintenant on trace le gradient !
    Call GradientFillRect(Me.hDC, VERT(0), 2, gRect, 1, Hconst)

End Sub

'=======================================================
'remplit une zone
'=======================================================
Public Sub FillArea(Left As Long, Top As Long, Width As Long, Height As _
    Long, Optional ByVal Color As Long)
Dim hBrush2 As Long
Dim tRect As RECT
    
    'on créé un brush de couleur
    hBrush2 = CreateSolidBrush(GetColor(Color))
    
    'on définit la zone
    With tRect
        .Left = Left
        .Top = Top
        .Right = Width
        .Bottom = Height
    End With

    'on remplit la zone et on libère le brush
    Call FillRect(Me.hDC, tRect, hBrush2)
    Call DeleteObject(hBrush2)
    
    'on remet l'ancien Brush si existant
    If hBrush Then Call SelectObject(Me.hDC, hBrush)

End Sub

'=======================================================
'inverse une zone
'=======================================================
Public Sub InvertArea(Left As Long, Top As Long, Width As Long, Height As _
    Long)
    
Dim tRect As RECT

    'définit le rectangle
    With tRect
        .Left = Left
        .Top = Top
        .Right = Width + Left
        .Bottom = Height + Top
    End With
    
    'inverse les couleurs
    Call InvertRect(lhDc, tRect)
End Sub

'=======================================================
'créé notre VirtualDC depuis une picturebox
'=======================================================
Public Sub CreateFromPictureBox(pct As PictureBox)
Dim hDesktop As Long
Dim lBmpCopy As Long
Dim hTempDC  As Long
Dim hMem As Long
    
    'récupère la police
    Set tFont = pct.Font
    
    'récupère les couleurs et les tailles
    With Me
        .BackColor = pct.BackColor
        .ForeColor = pct.ForeColor
        .Width = pct.Width / Screen.TwipsPerPixelX
        .Height = pct.Height / Screen.TwipsPerPixelY
    End With
    
    'récupère l'image
    Set Me.Picture = pct.Picture
    
    'récupère bureau
    hDesktop = GetDC(HWND_DESKTOP)

    'création DC compatible
    hTempDC = CreateCompatibleDC(hDesktop)

    'copie image
    lBmpCopy = CopyImage(pct.Picture.Handle, 0&, lWidth, lHeight, 0&)

    'on lie la copie l'image
    hMem = SelectObject(hTempDC, lBmpCopy)

    'libère le bureau
    Call ReleaseDC(HWND_DESKTOP, hDesktop)

    'création du DC compatible pour la classe
    Call Me.CreateMyDC(lWidth, lHeight)

    'copie l'image du pct
    Call BitBlt(Me.hDC, 0&, 0&, lWidth, lHeight, pct.hDC, 0&, 0&, vbSrcCopy)

    'libère
    Call SelectObject(hTempDC, hMem)
    Call DeleteDC(hTempDC)
    Call ReleaseDC(HWND_DESKTOP, hDesktop)
    
End Sub

'=======================================================
'efface tout
'=======================================================
Public Sub Cls()
    Call Me.FillArea(0, 0, lWidth, lHeight, vbWhite)
End Sub

'=======================================================
'ajoute un point de couleur
'=======================================================
Public Sub SetPixel(x As Long, y As Long, Color As Long)
    Call SetPixelV(Me.hDC, x, y, Color)
End Sub

'=======================================================
'récupère la couleur d'ub point
'=======================================================
Public Function GetPixel(x As Long, y As Long) As Long
    GetPixel = GetPixelAPI(Me.hDC, x, y)
End Function

'=======================================================
'dessine un rectangle de focus
'=======================================================
Public Sub DrawFocusRect(Left As Long, Top As Long, Width As Long, _
    Height As Long)
    
Dim R As RECT
    
    'définit les rectangles
    With R
        .Left = Left
        .Top = Top
        .Right = Width + Left
        .Bottom = Height + Top
    End With
    
    'dessine le rectangle
    Call DrawFocusRectAPI(Me.hDC, R)
End Sub

'=======================================================
'dessine une ellipse
'=======================================================
Public Sub DrawEllipse(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
    
    'trace l'ellipse
    Call Ellipse(Me.hDC, X1, Y1, X2, Y2)

End Sub

'=======================================================
'dessine un rectangle
'=======================================================
Public Sub DrawRectangle(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
    
    'trace le rectangle
    Call Rectangle(Me.hDC, X1, Y1, X2, Y2)
    
End Sub

'=======================================================
'dessine un rectangle a bords arrondis
'=======================================================
Public Sub DrawRoundRectangle(X1 As Long, Y1 As Long, X2 As Long, _
    Y2 As Long, X3 As Long, Y3 As Long)

    'trace le rectangle
    Call RoundRect(Me.hDC, X1, Y1, X2, Y2, X3, Y3)

End Sub

'=======================================================
'dessine un "edge"
'=======================================================
Public Sub DrawEdge(Left As Long, Top As Long, Width As Long, _
    Height As Long, Edge As EdgeEnum, Flags As EdgeFlags)
    
Dim R As RECT

    'définit la zone
    With R
        .Left = Left
        .Top = Top
        .Right = Width + Left
        .Bottom = Height + Top
    End With
    
    'trace
    Call DrawEdgeAPI(Me.hDC, R, Edge, Flags)
End Sub

'=======================================================
'détruit le DC et tous les objets
'=======================================================
Public Sub DestroyDC()
    Call Class_Terminate
End Sub



'=======================================================
'PRIVATE SUB/FUNCTIONS
'=======================================================
Private Function hDCToPicture(ByVal hdcSrc As Long, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
    Dim hDCMemory As Long, hBmp As Long, hBmpPrev As Long, R As Long
    Dim hPal As Long, hPalPrev As Long, RasterCapsScrn As Long, HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long, LogPal As LOGPALETTE
    
    hDCMemory = CreateCompatibleDC(hdcSrc)
    hBmp = CreateCompatibleBitmap(hdcSrc, WidthSrc, HeightSrc)
    hBmpPrev = SelectObject(hDCMemory, hBmp)
    
    RasterCapsScrn = GetDeviceCaps(hdcSrc, RASTERCAPS)   ' Raster
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette
    PaletteSizeScrn = GetDeviceCaps(hdcSrc, SIZEPALETTE) ' Size of

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        R = GetSystemPaletteEntries(hdcSrc, 0, 256, LogPal.palPalEntry(0))
        hPal = CreatePalette(LogPal)
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        R = RealizePalette(hDCMemory)
    End If

    R = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hdcSrc, LeftSrc, TopSrc, vbSrcCopy)
    
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If
    
    R = DeleteDC(hDCMemory)
    
    Set hDCToPicture = CreateBitmapPicture(hBmp, hPal)
End Function
Private Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    Dim R As Long, pic As PicBmp, IPic As IPicture, IID_IDispatch As GUID

    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    
    With pic
        .Size = Len(pic)          ' Length of structure
        .Type = vbPicTypeBitmap   ' Type of Picture (bitmap)
        .hBmp = hBmp              ' Handle to bitmap
        .hPal = hPal              ' Handle to palette (may be null)
    End With
    
    R = OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)
    
    Set CreateBitmapPicture = IPic
End Function
'=======================================================
'récupère une couleur compatible
'=======================================================
Private Function GetColor(ByVal nColor As Long) As Long
    Const SYSCOLOR_BIT As Long = &H80000000
    If (nColor And SYSCOLOR_BIT) = SYSCOLOR_BIT Then
        nColor = nColor And (Not SYSCOLOR_BIT)
        GetColor = GetSysColor(nColor)
    Else
        GetColor = nColor
    End If
End Function
'=======================================================
'associe une fonte au DC
'=======================================================
Private Sub AssociateFont(Font As StdFont)
Dim nName() As Byte
Dim i As Byte
Dim tFont As LOGFONT
Dim nSize As Byte

    nName = StrConv(Font.Name & Chr$(0), vbFromUnicode)
    nSize = UBound(nName)
    If nSize > 32 Then nSize = 32
    For i = 0 To nSize
        tFont.lfFaceName(i) = nName(i)
    Next
    
    With tFont
        .lfCharSet = Font.Charset
        .lfEscapement = 0
        .lfOrientation = .lfEscapement
        .lfItalic = IIf(Font.Italic, 1, 0)
        .lfStrikeOut = Abs(CLng(Font.Strikethrough))
        .lfUnderline = Abs(CLng(Font.Underline))
        .lfWeight = Font.Weight
        .lfHeight = MulDiv(Font.Size, GetDeviceCaps(Me.hDC, _
            LOGPIXELSY), 72)
    End With
    
    'si déjà une fonte, alors on la supprime
    If hFont <> 0 Then
        Call SelectObject(Me.hDC, hOldFont)
        Call DeleteObject(hFont)
    End If
    
    'on sélectionne la nouvelle fonte
    hFont = CreateFontIndirect(tFont)
    hOldFont = SelectObject(Me.hDC, hFont)

End Sub
'=======================================================
'convertit une couleur en long vers RGB (Integer)
'=======================================================
Private Function LongToRGBint(ByVal Color As Long) As RGB_COLOR_INT
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
'sélectionne un nouveau Pen
'=======================================================
Private Sub SelectNewPen()
    
    'on vire l'ancien
    If hPen Then Call DeleteObject(hPen)
    
    'on créé le nouveau
    hPen = CreatePen(ttPenStyle, lPenWidth, lPenColor)
    
    'on l'associe
    Call SelectObject(Me.hDC, hPen)
    
End Sub
'=======================================================
'sélectionne un nouveau Brush
'=======================================================
Private Sub SelectNewBrush()
Dim LB As LOGBRUSH
    
    'on vire l'ancien
    If hBrush Then Call DeleteObject(hBrush)
    
    With LB
        .lbColor = lBrushColor
        .lbHatch = tBrushHatch
        .lbStyle = tBrushStyle
    End With
        
    
    'on créé le nouveau
    hBrush = CreateBrushIndirect(LB)
    
    'on l'associe
    Call SelectObject(Me.hDC, hBrush)
    
End Sub
