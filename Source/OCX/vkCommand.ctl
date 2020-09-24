VERSION 5.00
Begin VB.UserControl vkCommand 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   DefaultCancel   =   -1  'True
   PropertyPages   =   "vkCommand.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "vkCommand.ctx":004D
   Begin VB.PictureBox pctG 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1800
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image PCTmouse 
      Height          =   240
      Left            =   2640
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image PCTgray 
      Height          =   240
      Left            =   1920
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image PCTcolor 
      Height          =   240
      Left            =   2280
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "vkCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
Private mAsm(63) As Byte    'contient le code ASM
Private OldProc As Long     'adresse de l'ancienne window proc
Private objHwnd As Long     'handle de l'objet concerné
Private ET As TRACKMOUSEEVENTTYPE   'type pour le mouse_hover et le mouse_leave
Private IsMouseIn As Boolean    'si la souris est dans le controle

Private bPushed As Boolean
Private lTextPos As AlignmentConstants
Private lForeColor As OLE_COLOR
Private bCol1 As OLE_COLOR
Private bCol2 As OLE_COLOR
Private tCol1 As OLE_COLOR
Private tCol2 As OLE_COLOR
Private sCaption As String
Private lGradient As GradientConstants
Private bNotOk As Boolean
Private bNotOk2 As Boolean
Private bEnable As Boolean
Private lBorderColor As OLE_COLOR
Private bBreakCorner As Boolean
Private pctAlign As PictureAlignment
Private bPic As Boolean
Private lOffsetX As Long
Private lOffsetY As Long
Private bGray As Boolean
Private bDrawFocus As Boolean
Private bDrawMouseInRect As Boolean
Private bHasFocus As Boolean
Private lNotEnabledColor As OLE_COLOR
Private bUnRefreshControl As Boolean
Private bHasLeftOneTime As Boolean
Private tCustomStyle As vkCommandStyle
Private bJustPushedTab As Boolean
Private bUnicode As Boolean
Private bRefOneTime As Boolean
Private bDisplayMouseHoverIcon As Boolean


'=======================================================
'EVENTS
'=======================================================
Public Event Click()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseHover()
Public Event MouseLeave()
Public Event MouseWheel(Sens As Wheel_Sens)
Public Event MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
Public Event MouseUp(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
Public Event MouseDblClick(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
Public Event MouseMove(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)




'=======================================================
'USERCONTROL SUBS
'=======================================================
'=======================================================
' /!\ NE PAS DEPLACER CETTE FONCTION /!\ '
'=======================================================
' Cette fonction doit rester la premiere '
' fonction "public" du module de classe  '
'=======================================================
Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Attribute WindowProc.VB_MemberFlags = "40"
Dim iControl As Integer
Dim iShift As Integer
Dim z As Long
Dim x As Long
Dim e As Boolean
Dim y As Long

    Select Case uMsg
        
        Case WM_LBUTTONDBLCLK
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
                If bEnable Then _
                bPushed = True: Refresh
                RaiseEvent MouseDblClick(vbLeftButton, iShift, iControl, x, y)
        Case WM_LBUTTONDOWN
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
                If bEnable Then _
                bPushed = True: Refresh
                RaiseEvent MouseDown(vbLeftButton, iShift, iControl, x, y)
        Case WM_LBUTTONUP
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
                e = bPushed
                bPushed = False: Refresh
                Call DrawMouseEnterRect
                If e And bEnable Then RaiseEvent Click
                bPushed = False
                RaiseEvent MouseUp(vbLeftButton, iShift, iControl, x, y)
        Case WM_MBUTTONDBLCLK
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
                RaiseEvent MouseDblClick(vbMiddleButton, iShift, iControl, x, y)
        Case WM_MBUTTONDOWN
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
                RaiseEvent MouseDown(vbMiddleButton, iShift, iControl, x, y)
        Case WM_MBUTTONUP
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
                RaiseEvent MouseUp(vbMiddleButton, iShift, iControl, x, y)
        Case WM_MOUSEHOVER
            If IsMouseIn = False Then
                IsMouseIn = True
                If bRefOneTime = False And bDisplayMouseHoverIcon Then
                    bRefOneTime = True
                    Call Refresh(False)
                End If
                Call DrawMouseEnterRect
                RaiseEvent MouseHover
            End If
        Case WM_MOUSELEAVE
            RaiseEvent MouseLeave
            bPushed = False
            IsMouseIn = False
            bNotOk = False
            bRefOneTime = False
            If bHasLeftOneTime Then
                If bDrawMouseInRect And bDrawFocus Then _
                Call Refresh(False): Call DrawFocusRects
            Else
                bHasLeftOneTime = True
            End If
        Case WM_MOUSEMOVE
            Call TrackMouseEvent(ET)
            
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
    
                If (wParam And MK_LBUTTON) = MK_LBUTTON Then z = vbLeftButton
                If (wParam And MK_RBUTTON) = MK_RBUTTON Then z = vbRightButton
                If (wParam And MK_MBUTTON) = MK_MBUTTON Then z = vbMiddleButton
                RaiseEvent MouseMove(z, iShift, iControl, x, y)
                
                If z = vbLeftButton Then
                    If Not IsMouseInCtrl Then
                        If bPushed Then
                            bPushed = False
                            Call Refresh
                        End If
                    Else
                        If Not bPushed Then
                            bPushed = True
                            Call Refresh
                        End If
                    End If
                End If
        
        Case WM_RBUTTONDBLCLK
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
                RaiseEvent MouseDblClick(vbRightButton, iShift, iControl, x, y)
        Case WM_RBUTTONDOWN
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
                RaiseEvent MouseDown(vbRightButton, iShift, iControl, x, y)
        Case WM_RBUTTONUP
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
                RaiseEvent MouseUp(vbRightButton, iShift, iControl, x, y)
        Case WM_MOUSEWHEEL
            If wParam < 0 Then
                RaiseEvent MouseWheel(WHEEL_DOWN)
            Else
                RaiseEvent MouseWheel(WHEEL_UP)
            End If
        Case WM_PAINT
            bNotOk = True  'évite le clignotement lors du survol de la souris
    End Select
    
    'appel de la routine standard pour les autres messages
    WindowProc = CallWindowProc(OldProc, hWnd, uMsg, wParam, lParam)
    
End Function

Private Sub UserControl_ExitFocus()
'/!\ WARNING, DO NOT REMOVE
'IT PREVENT A BUG WHEN YOU ADD "FORM.SHOW VBMODAL" IN CLICK() EVENT
    Call SendMessage(UserControl.hWnd, WM_LBUTTONUP, 0, 0)
End Sub

Private Sub UserControl_Initialize()
Dim Ofs As Long
Dim Ptr As Long
        
    'Recupere l'adresse de "Me.WindowProc"
    Call CopyMemory(Ptr, ByVal (ObjPtr(Me)), 4)
    Call CopyMemory(Ptr, ByVal (Ptr + 489 * 4), 4)
    
    'Crée la veritable fonction WindowProc (à optimiser)
    Ofs = VarPtr(mAsm(0))
    MovL Ofs, &H424448B            '8B 44 24 04          mov         eax,dword ptr [esp+4]
    MovL Ofs, &H8245C8B            '8B 5C 24 08          mov         ebx,dword ptr [esp+8]
    MovL Ofs, &HC244C8B            '8B 4C 24 0C          mov         ecx,dword ptr [esp+0Ch]
    MovL Ofs, &H1024548B           '8B 54 24 10          mov         edx,dword ptr [esp+10h]
    MovB Ofs, &H68                 '68 44 33 22 11       push        Offset RetVal
    MovL Ofs, VarPtr(mAsm(59))
    MovB Ofs, &H52                 '52                   push        edx
    MovB Ofs, &H51                 '51                   push        ecx
    MovB Ofs, &H53                 '53                   push        ebx
    MovB Ofs, &H50                 '50                   push        eax
    MovB Ofs, &H68                 '68 44 33 22 11       push        ObjPtr(Me)
    MovL Ofs, ObjPtr(Me)
    MovB Ofs, &HE8                 'E8 1E 04 00 00       call        Me.WindowProc
    MovL Ofs, Ptr - Ofs - 4
    MovB Ofs, &HA1                 'A1 20 20 40 00       mov         eax,RetVal
    MovL Ofs, VarPtr(mAsm(59))
    MovL Ofs, &H10C2               'C2 10 00             ret         10h
End Sub

Private Sub UserControl_InitProperties()
    'valeurs par défaut
    bNotOk2 = True
    With Me
        .BackColor1 = &HFBFBFB       '
        .BackColor2 = &HDCDCDC    '
        .BackColorPushed1 = &HC8C8C8          '
        .BackColorPushed2 = &HEBEBEB       '
        .BackGradient = Horizontal '
        .Caption = "Caption" '
        .Font = Ambient.Font '
        .ForeColor = 7552000 '
        .TextPosition = vbCenter '
        .Enabled = True '
        .BorderColor = 7552000    '
        .BreakCorner = True '
        Set .Picture = Nothing
        .PictureAlignment = [Left Justify]
        .DisplayPicture = True
        .PictureOffsetX = 0
        .PictureOffsetY = 0
        .GrayPictureWhenDisabled = True
        .DrawFocus = True
        .DrawMouseInRect = True
        .DisabledBackColor = 15198183
        .UnRefreshControl = False
        .CustomStyle = [Defaut Style - vkCommand]
        .DisplayMouseHoverIcon = False
        Set .MouseHoverPicture = Nothing
        .UseUnicode = False
    End With
    bNotOk2 = False
    'Call UserControl_Paint  'refresh
End Sub

Private Sub UserControl_GotFocus()

     If bEnable = False Then
        'on ne garde pas le focus
        Call SendKeys("{Tab}")
        Exit Sub
    End If
    
    'on a alors le focus
    bHasFocus = True
    
    'trace le rectangle de focus
    Call DrawFocusRects
    
    'trace aussi le rectangle de survol si nécessaire
    If IsMouseInCtrl Then Call DrawMouseEnterRect
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    If bEnable Then RaiseEvent Click
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If bEnable = False Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyUp, vbKeyLeft:
            Call SendKeys("+{Tab}")
            GoTo 12
        Case vbKeyDown, vbKeyRight:
            Call SendKeys("{Tab}")
            GoTo 12
        Case vbKeySpace
            If bPushed = False Then
                bPushed = True: Refresh: RaiseEvent Click
            End If
        Case vbKeyReturn
            RaiseEvent Click
        Case vbKeyTab
            GoTo 12
    End Select
    
    'Call Refresh
    RaiseEvent KeyDown(KeyCode, Shift)
    
    Exit Sub
12:
'bJustPushedTab = True
'    bHasFocus = False
'    IsMouseIn = False
'    Refresh
'    Call SendMessage(UserControl.hWnd, WM_MOUSELEAVE, 0&, 0&)
'UserControl_LostFocus
'Refresh
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
    If KeyCode = vbKeySpace Then
        bPushed = False
        Call Refresh
        Call DrawMouseEnterRect
    End If
End Sub

Private Sub UserControl_LostFocus()
    bHasFocus = False
    bNotOk = False
    bJustPushedTab = True
    Call UserControl_Paint
End Sub

Private Sub UserControl_Terminate()
    'vire le subclassing
    If OldProc Then Call SetWindowLong(UserControl.hWnd, GWL_WNDPROC, OldProc)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("UnRefreshControl", Me.UnRefreshControl, False)
        Call .WriteProperty("BackColor1", Me.BackColor1, &HFBFBFB)
        Call .WriteProperty("BackColor2", Me.BackColor2, &HDCDCDC)
        Call .WriteProperty("BackColorPushed1", Me.BackColorPushed1, &HC8C8C8)
        Call .WriteProperty("BackColorPushed2", Me.BackColorPushed2, &HEBEBEB)
        Call .WriteProperty("BackGradient", Me.BackGradient, Horizontal)
        Call .WriteProperty("Caption", Me.Caption, "Caption")
        Call .WriteProperty("Font", Me.Font, Ambient.Font)
        Call .WriteProperty("ForeColor", Me.ForeColor, 7552000)
        Call .WriteProperty("TextPosition", Me.TextPosition, vbCenter)
        Call .WriteProperty("Enabled", Me.Enabled, True)
        Call .WriteProperty("BorderColor", Me.BorderColor, 7552000)
        Call .WriteProperty("BreakCorner", Me.BreakCorner, True)
        Call .WriteProperty("Picture", Me.Picture, Nothing)
        Call .WriteProperty("PictureAlignment", Me.PictureAlignment, [Left Justify])
        Call .WriteProperty("DisplayPicture", Me.DisplayPicture, True)
        Call .WriteProperty("PictureOffsetX", Me.PictureOffsetX, 0)
        Call .WriteProperty("PictureOffsetY", Me.PictureOffsetY, 0)
        Call .WriteProperty("GrayPictureWhenDisabled", Me.GrayPictureWhenDisabled, True)
        Call .WriteProperty("DrawFocus", Me.DrawFocus, True)
        Call .WriteProperty("DrawMouseInRect", Me.DrawMouseInRect, True)
        Call .WriteProperty("DisabledBackColor", Me.DisabledBackColor, 15198183)
        Call .WriteProperty("CustomStyle", Me.CustomStyle, [Defaut Style - vkCommand])
        Call .WriteProperty("UseUnicode", Me.UseUnicode, False)
        Call .WriteProperty("DisplayMouseHoverIcon", Me.DisplayMouseHoverIcon, False)
        Call .WriteProperty("MouseHoverPicture", Me.MouseHoverPicture, Nothing)
    End With
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    bNotOk2 = True
    With PropBag
        Me.BackColor1 = .ReadProperty("BackColor1", &HFBFBFB)
        Me.BackColor2 = .ReadProperty("BackColor2", &HDCDCDC)
        Me.BackColorPushed1 = .ReadProperty("BackColorPushed1", &HC8C8C8)
        Me.BackColorPushed2 = .ReadProperty("BackColorPushed2", &HEBEBEB)
        Me.BackGradient = .ReadProperty("BackGradient", Horizontal)
        Me.Caption = .ReadProperty("Caption", "Caption")
        Set Me.Font = .ReadProperty("Font", Ambient.Font)
        Me.ForeColor = .ReadProperty("ForeColor", 7552000)
        Me.TextPosition = .ReadProperty("TextPosition", vbCenter)
        Me.Enabled = .ReadProperty("Enabled", True)
        Me.BorderColor = .ReadProperty("BorderColor", 7552000)
        Me.BreakCorner = .ReadProperty("BreakCorner", True)
        Set Me.Picture = .ReadProperty("Picture", Nothing)
        Set Me.MouseHoverPicture = .ReadProperty("MouseHoverPicture", Nothing)
        Me.DisplayMouseHoverIcon = .ReadProperty("DisplayMouseHoverIcon", False)
        Me.PictureAlignment = .ReadProperty("PictureAlignment", [Left Justify])
        Me.DisplayPicture = .ReadProperty("DisplayPicture", True)
        Me.PictureOffsetX = .ReadProperty("PictureOffsetX", 0)
        Me.PictureOffsetY = .ReadProperty("PictureOffsetY", 0)
        Me.GrayPictureWhenDisabled = .ReadProperty("GrayPictureWhenDisabled", True)
        Me.DrawFocus = .ReadProperty("DrawFocus", True)
        Me.DrawMouseInRect = .ReadProperty("DrawMouseInRect", True)
        Me.DisabledBackColor = .ReadProperty("DisabledBackColor", 15198183)
        Me.UnRefreshControl = .ReadProperty("UnRefreshControl", False)
        Me.CustomStyle = .ReadProperty("CustomStyle", [Defaut Style - vkCommand])
        Me.UseUnicode = .ReadProperty("UseUnicode", False)
    End With
    bNotOk2 = False
    'Call UserControl_Paint  'refresh
    Call Refresh
    
    'le bon endroit pour lancer le subclassing
    Call LaunchKeyMouseEvents
End Sub
Private Sub UserControl_Resize()
    bNotOk = False
    Call UserControl_Paint  'refresh
End Sub

'=======================================================
'lance le subclassing
'=======================================================
Private Sub LaunchKeyMouseEvents()
                
    If Ambient.UserMode Then

        OldProc = SetWindowLong(UserControl.hWnd, GWL_WNDPROC, _
            VarPtr(mAsm(0)))    'pas de AddressOf aujourd'hui ;)
            
        'prépare le terrain pour le mouse_over et mouse_leave
        With ET
            .cbSize = Len(ET)
            .hwndTrack = UserControl.hWnd
            .dwFlags = TME_LEAVE Or TME_HOVER
            .dwHoverTime = 1
        End With
        
        'démarre le tracking de l'entrée
        Call TrackMouseEvent(ET)
        
        'pas dedans par défaut
        IsMouseIn = False
        
    End If
    
End Sub



'=======================================================
'PROPERTIES
'=======================================================
Public Property Get hDC() As Long: hDC = UserControl.hDC: End Property
Public Property Get hWnd() As Long: hWnd = UserControl.hWnd: End Property
Public Property Get TextPosition() As AlignmentConstants: TextPosition = lTextPos: End Property
Public Property Let TextPosition(TextPosition As AlignmentConstants): lTextPos = TextPosition: bNotOk = False: UserControl_Paint: End Property
Public Property Get Caption() As String: Caption = sCaption: End Property
Public Property Let Caption(Caption As String): sCaption = Caption: Call SetAccessKeys: bNotOk = False: UserControl_Paint: bNotOk = True: End Property
Public Property Get ForeColor() As OLE_COLOR: ForeColor = lForeColor: End Property
Public Property Let ForeColor(ForeColor As OLE_COLOR): lForeColor = ForeColor: UserControl.ForeColor = ForeColor: bNotOk = False: UserControl_Paint: tCustomStyle = [No Style - vkCommand]: End Property
Public Property Get BackColor1() As OLE_COLOR: BackColor1 = bCol1: End Property
Public Property Let BackColor1(BackColor1 As OLE_COLOR): tCustomStyle = [No Style - vkCommand]: bCol1 = BackColor1: bNotOk = False: UserControl_Paint: End Property
Public Property Get BackColor2() As OLE_COLOR: BackColor2 = bCol2: End Property
Public Property Let BackColor2(BackColor2 As OLE_COLOR): tCustomStyle = [No Style - vkCommand]: bCol2 = BackColor2: bNotOk = False: UserControl_Paint: End Property
Public Property Get BackColorPushed1() As OLE_COLOR: BackColorPushed1 = tCol1: End Property
Public Property Let BackColorPushed1(BackColorPushed1 As OLE_COLOR): tCustomStyle = [No Style - vkCommand]: tCol1 = BackColorPushed1: bNotOk = False: UserControl_Paint: End Property
Public Property Get BackColorPushed2() As OLE_COLOR: BackColorPushed2 = tCol2: End Property
Public Property Let BackColorPushed2(BackColorPushed2 As OLE_COLOR): tCustomStyle = [No Style - vkCommand]: tCol2 = BackColorPushed2: bNotOk = False: UserControl_Paint: End Property
Public Property Get Font() As StdFont: Set Font = UserControl.Font: End Property
Public Property Set Font(Font As StdFont): Set UserControl.Font = Font: bNotOk = False: UserControl_Paint: End Property
Public Property Get BackGradient() As GradientConstants: BackGradient = lGradient: End Property
Public Property Let BackGradient(BackGradient As GradientConstants): tCustomStyle = [No Style - vkCommand]: lGradient = BackGradient: bNotOk = False: UserControl_Paint: End Property
Public Property Get Enabled() As Boolean: Enabled = bEnable: End Property
Public Property Let Enabled(Enabled As Boolean)
bEnable = Enabled: bNotOk = False: UserControl_Paint
End Property
Public Property Get BorderColor() As OLE_COLOR: BorderColor = lBorderColor: End Property
Public Property Let BorderColor(BorderColor As OLE_COLOR): tCustomStyle = [No Style - vkCommand]: lBorderColor = BorderColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get BreakCorner() As Boolean: BreakCorner = bBreakCorner: End Property
Public Property Let BreakCorner(BreakCorner As Boolean): tCustomStyle = [No Style - vkCommand]: bBreakCorner = BreakCorner: bNotOk = False: UserControl_Paint: End Property
Public Property Get Picture() As Picture: Set Picture = PCTcolor.Picture: End Property
Public Property Set Picture(NewPic As Picture)
Set PCTcolor.Picture = NewPic
Set pctG.Picture = NewPic
If Not (NewPic Is Nothing) Then
    pctG.Width = PCTcolor.Width
    pctG.Height = PCTcolor.Height
    Call GrayScale(pctG)
End If
PCTgray.Picture = pctG.Image
bNotOk = False: UserControl_Paint
End Property
Public Property Get PictureAlignment() As PictureAlignment: PictureAlignment = pctAlign: End Property
Public Property Let PictureAlignment(PictureAlignment As PictureAlignment): pctAlign = PictureAlignment: bNotOk = False: UserControl_Paint: End Property
Public Property Get DisplayPicture() As Boolean: DisplayPicture = bPic: End Property
Public Property Let DisplayPicture(DisplayPicture As Boolean): bPic = DisplayPicture: bNotOk = False: UserControl_Paint: End Property
Public Property Get PictureOffsetX() As Long: PictureOffsetX = lOffsetX: End Property
Public Property Let PictureOffsetX(PictureOffsetX As Long): lOffsetX = PictureOffsetX: bNotOk = False: UserControl_Paint: End Property
Public Property Get PictureOffsetY() As Long: PictureOffsetY = lOffsetY: End Property
Public Property Let PictureOffsetY(PictureOffsetY As Long): lOffsetY = PictureOffsetY: bNotOk = False: UserControl_Paint: End Property
Public Property Get GrayPictureWhenDisabled() As Boolean: GrayPictureWhenDisabled = bGray: End Property
Public Property Let GrayPictureWhenDisabled(GrayPictureWhenDisabled As Boolean): bGray = GrayPictureWhenDisabled: bNotOk = False: UserControl_Paint: End Property
Public Property Get DrawFocus() As Boolean: DrawFocus = bDrawFocus: End Property
Public Property Let DrawFocus(DrawFocus As Boolean): tCustomStyle = [No Style - vkCommand]: bDrawFocus = DrawFocus: bNotOk = False: UserControl_Paint: End Property
Public Property Get DrawMouseInRect() As Boolean: DrawMouseInRect = bDrawMouseInRect: End Property
Public Property Let DrawMouseInRect(DrawMouseInRect As Boolean): tCustomStyle = [No Style - vkCommand]: bDrawMouseInRect = DrawMouseInRect: bNotOk = False: UserControl_Paint: End Property
Public Property Get DisabledBackColor() As OLE_COLOR: DisabledBackColor = lNotEnabledColor: End Property
Public Property Let DisabledBackColor(DisabledBackColor As OLE_COLOR): tCustomStyle = [No Style - vkCommand]: lNotEnabledColor = DisabledBackColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get UnRefreshControl() As Boolean: UnRefreshControl = bUnRefreshControl: End Property
Public Property Let UnRefreshControl(UnRefreshControl As Boolean): bUnRefreshControl = UnRefreshControl: End Property
Public Property Get CustomStyle() As vkCommandStyle: CustomStyle = tCustomStyle: End Property
Public Property Let CustomStyle(CustomStyle As vkCommandStyle)
    tCustomStyle = CustomStyle
    
    Select Case CustomStyle
        Case [Defaut Style - vkCommand]
            bCol1 = &HFBFBFB
            bCol2 = &HDCDCDC
            tCol1 = &HC8C8C8
            tCol2 = &HEBEBEB
            lGradient = Horizontal
            lBorderColor = &H733C00
            bBreakCorner = True
            lNotEnabledColor = &HE7E7E7
            bDrawFocus = True
            bDrawMouseInRect = True
            lForeColor = &H733C00
        Case [Office Style - vkCommand]
            bCol1 = &HFFC0C0
            tCol1 = &HFFC0C0
            lGradient = None
            lBorderColor = &H808080
            bBreakCorner = True
            lNotEnabledColor = &HE7E7E7
            bDrawFocus = False
            bDrawMouseInRect = False
            lForeColor = &H404040
        Case [XP Style - vkCommand]
            bCol1 = &HFFFFFF
            bCol2 = &HDAE3E6
            tCol1 = &HDDE4E5
            tCol2 = &HDAE4E2
            lGradient = Horizontal
            lBorderColor = &H743C00
            bBreakCorner = True
            lNotEnabledColor = 15398133
            bDrawFocus = True
            bDrawMouseInRect = True
            lForeColor = &H404040
        Case [Plastik Style - vkCommand]
            bCol1 = &HEDFEFF
            bCol2 = &HD2E3E6
            tCol1 = &HC0D1D4
            tCol2 = &HCEDFE2
            lGradient = Horizontal
            lBorderColor = &H809494
            bBreakCorner = True
            lNotEnabledColor = &HE9F5F6
            bDrawFocus = False
            bDrawMouseInRect = False
            lForeColor = &H404040
        Case [Galaxy Style - vkCommand]
            bCol1 = &HFFFFFF
            bCol2 = &HC9DADD
            tCol1 = &HD8E9EC
            tCol2 = &HFFFFFF
            lGradient = Horizontal
            lBorderColor = &HA8B9BC
            bBreakCorner = True
            lNotEnabledColor = &HE5F3F4
            bDrawFocus = False
            bDrawMouseInRect = False
            lForeColor = &H404040
        Case [JAVA Style - vkCommand]
            bCol1 = &HD8E9EC
            tCol1 = &HC9DADD
            lGradient = None
            lBorderColor = &H99A8AC
            bBreakCorner = False
            lNotEnabledColor = &HD8E9EC
            bDrawFocus = False
            bDrawMouseInRect = False
            lForeColor = &H404040
    End Select
    
    bNotOk = False: Call UserControl_Paint
    
End Property
Public Property Get UseUnicode() As Boolean: UseUnicode = bUnicode: End Property
Public Property Let UseUnicode(UseUnicode As Boolean): bUnicode = UseUnicode: bNotOk = False: UserControl_Paint: bNotOk = True: End Property
Public Property Get MouseHoverPicture() As Picture: Set MouseHoverPicture = PCTmouse.Picture: End Property
Public Property Set MouseHoverPicture(NewPic As Picture)
Set PCTmouse.Picture = NewPic
bNotOk = False: UserControl_Paint
End Property
Public Property Get DisplayMouseHoverIcon() As Boolean: DisplayMouseHoverIcon = bDisplayMouseHoverIcon: End Property
Public Property Let DisplayMouseHoverIcon(DisplayMouseHoverIcon As Boolean): bDisplayMouseHoverIcon = DisplayMouseHoverIcon: bNotOk = False: End Property


Private Sub UserControl_Paint()

    If bNotOk Or bNotOk2 Then Exit Sub     'pas prêt à peindre
    
    Call Refresh    'on refresh
End Sub




'=======================================================
'PRIVATE SUBS
'=======================================================
'=======================================================
'copie un "byte"
'=======================================================
Private Sub MovB(Ofs As Long, ByVal Value As Long)
    Call CopyMemory(ByVal Ofs, Value, 1): Ofs = Ofs + 1
End Sub

'=======================================================
'copie un "long"
'=======================================================
Private Sub MovL(Ofs As Long, ByVal Value As Long)
    Call CopyMemory(ByVal Ofs, Value, 4): Ofs = Ofs + 4
End Sub

'=======================================================
'convertit une couleur en long vers RGB
'=======================================================
Private Sub ToRGB(ByVal Color As Long, ByRef RGB As RGB_COLOR)
    With RGB
        .R = Color And &HFF&
        .G = (Color And &HFF00&) \ &H100&
        .B = Color \ &H10000
    End With
End Sub

'=======================================================
'récupère la hauteur d'un caractère
'=======================================================
Private Function GetCharHeight() As Long
Dim Res As Long
    Res = GetTabbedTextExtent(UserControl.hDC, "A", 1, 0, 0)
    GetCharHeight = (Res And &HFFFF0000) \ &H10000
End Function


'=======================================================
'PUBLIC SUB
'=======================================================
'=======================================================
'on dessine tout
'=======================================================
Public Sub Refresh(Optional ByVal ShowFocusRects As Boolean = True)
Dim x As Long
Dim R As RECT
Dim hBrush As Long
Dim Rec As Long
Dim hRgn As Long
Dim W As Long
Dim H As Long
    
    If bUnRefreshControl Then Exit Sub
    
    '//on efface et on vire le maskpicture
    Call UserControl.Cls
    UserControl.Picture = Nothing
    UserControl.MaskPicture = Nothing
    
    '//on convertir les différentes couleurs si couleurs système
    Call OleTranslateColor(lBorderColor, 0, lBorderColor)
    Call OleTranslateColor(lForeColor, 0, lForeColor)
    Call OleTranslateColor(bCol1, 0, bCol1)
    Call OleTranslateColor(bCol2, 0, bCol2)
    
    
    '//on va tracer le rectangle de focus si on a le focus
    If bHasFocus And ShowFocusRects Then Call UserControl_GotFocus
    
    '//créé le gradient du contrôle
    If bEnable Then
        
        If bPushed = False Then
            'pas appuyé
            
            'récupère les 3 composantes des deux couleurs
            Call OleTranslateColor(bCol1, 0, bCol1)
            Call OleTranslateColor(bCol2, 0, bCol2)
            
            If lGradient = None Then
                'pas de gradient
                'on dessine alors un rectangle
                Line (15, 0)-(ScaleWidth - 30, _
                    ScaleHeight - 30), bCol1, BF
            ElseIf lGradient = Horizontal Then
                'gradient horizontal
                Call FillGradient(UserControl.hDC, bCol1, bCol2, _
                    Width / Screen.TwipsPerPixelX, Height / _
                    Screen.TwipsPerPixelY, Horizontal)
            Else
                'gradient vertical
                Call FillGradient(UserControl.hDC, bCol1, bCol2, _
                    Width / Screen.TwipsPerPixelX, Height / _
                    Screen.TwipsPerPixelY, Vertical)
            End If
        Else
            'appuyé sur le bouton
            
            'récupère les 3 composantes des deux couleurs
            Call OleTranslateColor(tCol1, 0, tCol1)
            Call OleTranslateColor(tCol2, 0, tCol2)
            
            If lGradient = None Then
                'pas de gradient
                'on dessine alors un rectangle
                Line (15, 0)-(ScaleWidth - 30, _
                    ScaleHeight - 30), bCol2, BF
            ElseIf lGradient = Horizontal Then
                'gradient horizontal
                Call FillGradient(UserControl.hDC, tCol1, tCol2, _
                    Width / Screen.TwipsPerPixelX, Height / _
                    Screen.TwipsPerPixelY, Horizontal)
            Else
                'gradient vertical
                Call FillGradient(UserControl.hDC, tCol1, tCol2, _
                    Width / Screen.TwipsPerPixelX, Height / _
                    Screen.TwipsPerPixelY, Vertical)
            End If
        End If
    Else
        'controle non enabled
        UserControl.BackColor = lNotEnabledColor
    End If
    

    '//créé un rectangle
    'calcule le nombre de lignes à afficher
    x = Int((TextWidth(sCaption) + 120) / Width) + 1
    
    If bPic And PCTcolor.Picture Then
        'alors on affiche la picture ==> décale le texte
        
        If pctAlign = [Left Justify] Then
            Call SetRect(R, 8 + PCTcolor.Width / Screen.TwipsPerPixelX / 2, (Height / Screen.TwipsPerPixelY - GetCharHeight * x) / 2, Width / Screen.TwipsPerPixelX - 4, Height / Screen.TwipsPerPixelY)
        Else
            Call SetRect(R, 4, (Height / Screen.TwipsPerPixelY - GetCharHeight * x) / 2, Width / Screen.TwipsPerPixelX - 8 - PCTcolor.Width / Screen.TwipsPerPixelX / 2, Height / Screen.TwipsPerPixelY)
        End If
    Else
        'pas de picture ==> pas de décalage
        Call SetRect(R, 4, (Height / Screen.TwipsPerPixelY - GetCharHeight * x) / 2, Width / Screen.TwipsPerPixelX - 4, Height / Screen.TwipsPerPixelY)
    End If
    
    'on affiche le caption
    If bEnable Then
        UserControl.ForeColor = lForeColor
    Else
        UserControl.ForeColor = 9934743
    End If
    
    If bUnicode = False Then
        If lTextPos = vbCenter Then
            'au centre
            Call DrawText(UserControl.hDC, sCaption, Len(sCaption), R, aDT_WORDBREAK + DT_CENTER)
        ElseIf lTextPos = vbRightJustify Then
            'à droite
            Call DrawText(UserControl.hDC, sCaption, Len(sCaption), R, aDT_WORDBREAK + DT_RIGHT)
        Else
            'à gauche
            Call DrawText(UserControl.hDC, sCaption, Len(sCaption), R, aDT_WORDBREAK + DT_LEFT)
        End If
    Else
        If lTextPos = vbCenter Then
            'au centre
            Call DrawTextW(UserControl.hDC, StrPtr(sCaption), Len(sCaption), R, aDT_WORDBREAK + DT_CENTER)
        ElseIf lTextPos = vbRightJustify Then
            'à droite
            Call DrawTextW(UserControl.hDC, StrPtr(sCaption), Len(sCaption), R, aDT_WORDBREAK + DT_RIGHT)
        Else
            'à gauche
            Call DrawTextW(UserControl.hDC, StrPtr(sCaption), Len(sCaption), R, aDT_WORDBREAK + DT_LEFT)
        End If
    End If
    
    
    
    '//on va se tracer la bitmap maintenant ^^
    If bPic And PCTcolor.Picture Then
        'on montre l'image présente
        
        'on choisit celle à afficher (Grise ou pas)
        If bEnable Or Not (bGray) Then
            If IsMouseIn Then
                With PCTmouse
                    'on move IMG en fonction du choix de Aligment
                    'on centre l'image dans la barre de titre
                    If pctAlign = [Left Justify] Then
                        'à gauche
                        
                        W = (Height - .Height) / 2 + lOffsetY
                        H = (Width - TextWidth(sCaption)) / 2 - .Width + lOffsetX - 4
                        If W < 0 Then W = 0
                        If H < 0 Then H = 0
                        
                        .Top = W
                        .Left = H '2 pxls + arrondi
                        
                    Else
                        'à droite
                        
                        W = (Height - .Height) / 2 + lOffsetY
                        H = (Width - TextWidth(sCaption)) / 2 + TextWidth(sCaption) / 2 + .Width + lOffsetX + 4
                        If H < 0 Then H = 0
                        If W < 0 Then W = 0
            
                        .Top = W
                        .Left = H '2 pxls + arrondi
                        
                    End If
                        
                    .Visible = True
                    PCTgray.Visible = False
                    PCTcolor.Visible = False
                End With
            Else
                With PCTcolor
                    'on move IMG en fonction du choix de Aligment
                    'on centre l'image dans la barre de titre
                    If pctAlign = [Left Justify] Then
                        'à gauche
                        
                        W = (Height - .Height) / 2 + lOffsetY
                        H = (Width - TextWidth(sCaption)) / 2 - .Width + lOffsetX - 4
                        If W < 0 Then W = 0
                        If H < 0 Then H = 0
                        
                        .Top = W
                        .Left = H '2 pxls + arrondi
                        
                    Else
                        'à droite
                        
                        W = (Height - .Height) / 2 + lOffsetY
                        H = (Width - TextWidth(sCaption)) / 2 + TextWidth(sCaption) / 2 + .Width + lOffsetX + 4
                        If H < 0 Then H = 0
                        If W < 0 Then W = 0
            
                        .Top = W
                        .Left = H '2 pxls + arrondi
                        
                    End If
                        
                    .Visible = True
                    PCTgray.Visible = False
                    PCTmouse.Visible = False
                End With
            End If
        Else
            With PCTgray
                'on move IMG en fonction du choix de Aligment
                'on centre l'image dans la barre de titre
                If pctAlign = [Left Justify] Then
                    'à gauche
                    
                    W = (Height - .Height) / 2 + lOffsetY
                    H = (Width - TextWidth(sCaption)) / 2 - .Width + lOffsetX
                    If W < 0 Then W = 0
                    If H < 0 Then H = 0
                    
                    .Top = W
                    .Left = H '2 pxls + arrondi
                    
                Else
                    'à droite
                    
                    W = (Height - .Height) / 2 + lOffsetY
                    H = (Width - TextWidth(sCaption)) / 2 + .Width + lOffsetX
                    If H < 0 Then H = 0
                    If W < 0 Then W = 0
        
                    .Top = W
                    .Left = H '2 pxls + arrondi
                    
                End If
                    
                .Visible = True
                PCTcolor.Visible = False
            End With
        End If
    Else
        'on masque l'image
        PCTgray.Visible = False
        PCTcolor.Visible = False
    End If
    
    
    
    '//dessine le contour
    If bBreakCorner = False Then
        'alors c'est un rectangle
        
        'on défini un brush
        hBrush = CreateSolidBrush(lBorderColor)
        
        'on définit une zone rectangulaire à bords arrondi
        hRgn = CreateRectRgn(0, 0, ScaleWidth / Screen.TwipsPerPixelX, _
            ScaleHeight / Screen.TwipsPerPixelY)
        
        'on dessine le contour
        Call FrameRgn(UserControl.hDC, hRgn, hBrush, 1, 1)

        'on détruit le brush et la zone
        Call DeleteObject(hBrush)
        Call DeleteObject(hRgn)
    
    Else
        'alors c'est un arrondi
        
        'on défini un brush
        hBrush = CreateSolidBrush(lBorderColor)
        
        'on définit une zone rectangulaire à bords arrondi
        hRgn = CreateRoundRectRgn(0, 0, ScaleWidth / Screen.TwipsPerPixelX, _
            ScaleHeight / Screen.TwipsPerPixelY, 7, 7)
        
        'on dessine le contour
        Call FrameRgn(UserControl.hDC, hRgn, hBrush, 1, 1)
        
        'on défini la zone rectangulaire arrondi comme nouvelle fenêtre
        Call SetWindowRgn(UserControl.hWnd, hRgn, True)

        'on détruit le brush et la zone
        Call DeleteObject(hBrush)
        Call DeleteObject(hRgn)
        
    End If

    
    bNotOk = True

    '//on délocke le controle --> a évité les clignotements
   ' Call LockWindowUpdate(0)
    
    'permet de refresh la bordure
    Call UserControl.Refresh
End Sub

'=======================================================
'trace le rectangle d'entrée sur le controle
'=======================================================
Private Sub DrawMouseEnterRect()
Dim x As Long
Dim W As Long
Dim H As Long
    
    'couleur orange foncée : 38631 et 3257087
    'couleur orange clair : 7064575
    'couleur orange très clair : 9033981,9889535
    
    'évite de dessiner le rectangle de focus si on a appuyé sur Tab pour
    'enlever le focus
    If bJustPushedTab Then
        bJustPushedTab = False
        Exit Sub
    End If
    
    If bDrawMouseInRect = False Or bEnable = False Then Exit Sub
    
    With UserControl
        H = .Height
        W = .Width
    End With
        
    '//trace les lignes orange foncé
    UserControl.ForeColor = 3257087
    Line (15, 60)-(15, H - 60)
    Line (W - 45, 60)-(W - 45, H - 60)
    Line (30, H - 60)-(W - 45, H - 60)
    Line (45, H - 45)-(W - 60, H - 45), 38631
    
    
    '//trace les lignes orange clair
    UserControl.ForeColor = 7064575
    Line (30, 30)-(30, H - 60)
    Line (W - 60, 30)-(W - 60, H - 60)
        

    '//trace la ligne orange très claire
    Line (30, 30)-(W - 45, 30), 9033981
    Line (60, 15)-(W - 75, 15), 9889535
    
    Call UserControl.Refresh

End Sub

'=======================================================
'trace le rectangle de focus
'=======================================================
Private Sub DrawFocusRects()
Dim x As Long
Dim W As Long
Dim H As Long
    
    'couleur bleu foncée : 15696491 et 15183500
    'couleur bleu clair : 15782325
    'couleur bleu très clair : 16242621,16771022
    
    If bDrawFocus = False Or IsMouseIn Or bHasFocus = False Or bEnable = False Then Exit Sub
    
    With UserControl
        H = .Height
        W = .Width
    End With
        
    '//trace les lignes bleu foncé
    UserControl.ForeColor = 15183500
    Line (15, 60)-(15, H - 60)
    Line (W - 45, 60)-(W - 45, H - 60)
    Line (30, H - 60)-(W - 45, H - 60)
    Line (45, H - 45)-(W - 60, H - 45), 15696491
    
    
    '//trace les lignes bleu clair
    UserControl.ForeColor = 15782325
    Line (30, 30)-(30, H - 60)
    Line (W - 60, 30)-(W - 60, H - 60)
        

    '//trace la ligne bleu très claire
    Line (30, 30)-(W - 45, 30), 16242621
    Line (60, 15)-(W - 75, 15), 16771022
    
    Call UserControl.Refresh

End Sub

'=======================================================
'renvoie l'objet extender de ce usercontrol (pour les propertypages)
'=======================================================
Friend Property Get MyExtender() As Object
    Set MyExtender = UserControl.Extender
End Property
Friend Property Let MyExtender(MyExtender As Object)
    Set UserControl.Extender = MyExtender
End Property

'=======================================================
'renvoie True si la souris est DANS le controle
'=======================================================
Private Function IsMouseInCtrl() As Boolean
Dim tMouse As POINTAPI
Dim R As RECT

    'récupère la position de la souris
    Call GetCursorPos(tMouse)
    
    'récupère un RECT de la position du controle
    Call GetWindowRect(UserControl.hWnd, R)
    
    'renvoie le résultat
    IsMouseInCtrl = (PtInRect(R, tMouse.x, tMouse.y) <> 0)
End Function

'=======================================================
'défini les touches de raccourci (avec '&')
'=======================================================
Private Sub SetAccessKeys()
Dim A As Long

    'récupère la position du '&' en partant de la fin
    A = InStrRev(sCaption, "&")
    
    'si le '&' existe et n'est pas tout à la fin
    If A <> Len(sCaption) And A <> 0 Then
        'on récupère le caractère qui est juste après
        AccessKeys = Mid$(sCaption, A + 1, 1)
    End If
End Sub
