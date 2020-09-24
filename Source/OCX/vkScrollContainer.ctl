VERSION 5.00
Begin VB.UserControl vkScrollContainer 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "vkScrollContainer.ctx":0000
   Begin vkUserContolsXP.vkHScrollPrivate HS 
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   2880
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
   End
   Begin vkUserContolsXP.vkVScrollPrivate VS 
      Height          =   2655
      Left            =   3960
      TabIndex        =   0
      Top             =   480
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   4683
   End
End
Attribute VB_Name = "vkScrollContainer"
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

Private lBackStyle As BackStyleConstants
Private bCol1 As OLE_COLOR
Private bCol2 As OLE_COLOR
Private bShowBackGround As Boolean
Private lBackGradient As GradientConstants
Private bNotOk As Boolean
Private bNotOk2 As Boolean
Private bEnable As Boolean
Private lBorderColor As OLE_COLOR
Private bDisplayBorder As Boolean
Private lBWidth As Long
Private bUnRefreshControl As Boolean
Private lX As Long
Private lY As Long
Private lAreaHeight As Long
Private lAreaWidth As Long
Private OldValueH As Long
Private OldValueV As Long
Private tScroll As New vkPrivateScroll
Private bNotResize As Boolean


'=======================================================
'EVENTS
'=======================================================
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
Public Event HScrollChange(Value As Currency)
Public Event HScrollScroll()
Public Event VScrollChange(Value As Currency)
Public Event VScrollScroll()



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
Dim y As Long
   
    Select Case uMsg
        
        Case WM_LBUTTONDBLCLK
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
            RaiseEvent MouseDblClick(vbLeftButton, iShift, iControl, x, y)
        Case WM_LBUTTONDOWN
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
                RaiseEvent MouseDown(vbLeftButton, iShift, iControl, x, y)
        Case WM_LBUTTONUP
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
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
                RaiseEvent MouseHover
                IsMouseIn = True
            End If
        Case WM_MOUSELEAVE
            RaiseEvent MouseLeave
            IsMouseIn = False
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

Private Sub UserControl_Initialize()
Dim Ofs As Long
Dim Ptr As Long
    
    'on initialise le tScroll
    With tScroll
        .SmallChange = 100
        .LargeChange = 600
        .WheelChange = 300
        .Enabled = True
        .Value = 0
    End With
    
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
    
    'initialise leS VS
    bUnRefreshControl = True
    With VS
        .UnRefreshControl = True
        .Max = 2
        .Min = 1
        .Value = 1
        .Visible = True
        .Enabled = True
        .SmallChange = 100
        .LargeChange = 600
        .WheelChange = 300
        .UnRefreshControl = False
        'Call .Refresh
    End With
    With HS
        .UnRefreshControl = True
        .Max = 2
        .Min = 0
        .Value = 1
        .SmallChange = 100
        .LargeChange = 600
        .WheelChange = 300
        .Visible = True
        .Enabled = True
        .UnRefreshControl = False
        'Call .Refresh
    End With
    bUnRefreshControl = False
End Sub

Private Sub UserControl_InitProperties()
    'valeurs par défaut
    bNotOk2 = True
    With Me
        .BackColor1 = &HFBFBFB       '
        .BackColor2 = &HDCDCDC    '
        .BackGradient = Horizontal '
        .BackStyle = Opaque
        .ShowBackGround = True '
        .Enabled = True '
        .BorderColor = &HFF8080    '
        .DisplayBorder = True '
        .BorderWidth = 1 '
        .UnRefreshControl = False
        .x = 0
        .y = 0
        .AreaHeight = 0
        .AreaWidth = 0
    End With
    bNotOk2 = False
    Call UserControl_Paint  'refresh
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_Show()
Dim ctr As Control
    
    On Error Resume Next
    
    'on passe chaque controle en arrière plan
    For Each ctr In UserControl.ContainedControls
        Call ctr.ZOrder(vbSendToBack)
    Next ctr
    
    Call UserControl_Resize
    
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
        Call .WriteProperty("BackGradient", Me.BackGradient, Horizontal)
        Call .WriteProperty("BackStyle", Me.BackStyle, Opaque)
        Call .WriteProperty("ShowBackGround", Me.ShowBackGround, True)
        Call .WriteProperty("Enabled", Me.Enabled, True)
        Call .WriteProperty("BorderColor", Me.BorderColor, &HFF8080)
        Call .WriteProperty("DisplayBorder", Me.DisplayBorder, True)
        Call .WriteProperty("BorderWidth", Me.BorderWidth, 1)
        Call .WriteProperty("X", Me.x, 0)
        Call .WriteProperty("Y", Me.y, 0)
        Call .WriteProperty("AreaHeight", Me.AreaHeight, 0)
        Call .WriteProperty("AreaWidth", Me.AreaWidth, 0)
    End With
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    bNotOk2 = True
    With PropBag
        Me.BackColor1 = .ReadProperty("BackColor1", &HFBFBFB)
        Me.BackColor2 = .ReadProperty("BackColor2", &HDCDCDC)
        Me.BackGradient = .ReadProperty("BackGradient", Horizontal)
        Me.BackStyle = .ReadProperty("BackStyle", Opaque)
        Me.ShowBackGround = .ReadProperty("ShowBackGround", True)
        Me.Enabled = .ReadProperty("Enabled", True)
        Me.BorderColor = .ReadProperty("BorderColor", &HFF8080)
        Me.DisplayBorder = .ReadProperty("DisplayBorder", True)
        Me.BorderWidth = .ReadProperty("BorderWidth", 1)
        Me.UnRefreshControl = .ReadProperty("UnRefreshControl", False)
        Me.x = .ReadProperty("X", 0)
        Me.y = .ReadProperty("Y", 0)
        Me.AreaHeight = .ReadProperty("AreaHeight", 0)
        Me.AreaWidth = .ReadProperty("AreaWidth", 0)
        Me.VScroll = .ReadProperty("VScroll", tScroll)
        Me.HScroll = .ReadProperty("HScroll", tScroll)
    End With
    bNotOk2 = False
    'Call UserControl_Paint  'refresh
    Call Refresh
    
    'le bon endroit pour lancer le subclassing
    Call LaunchKeyMouseEvents
    
    If Ambient.UserMode Then
        Call VS.LaunchKeyMouseEvents    'subclasse également le VS
        Call HS.LaunchKeyMouseEvents    'et le HS
    End If
    
    UserControl_Show
End Sub
Private Sub UserControl_Resize()

    '//on change de place les VS et HS
    With VS
        .Top = 0 ' 15 * lBWidth + 15
        .Left = Width - .Width '- 15 * lBWidth - 15
        .Height = Height - IIf(Width < lAreaWidth, HS.Height, 0)
    End With
    With HS
        .Top = Height - .Height '- 15 * lBWidth - 15
        .Left = 0 ' 15 * lBWidth + 15
        .Width = Width - IIf(Height < lAreaHeight, VS.Width - 15, 0) '- 30 * lBWidth - 30
    End With
    
    If bNotResize = False Then Call ReorganizeView
    
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
Public Property Get BackStyle() As BackStyleConstants: BackStyle = lBackStyle: End Property
Public Property Let BackStyle(BackStyle As BackStyleConstants): lBackStyle = BackStyle: UserControl.BackStyle = BackStyle: bNotOk = False: UserControl_Paint: End Property
Public Property Get BackColor1() As OLE_COLOR: BackColor1 = bCol1: End Property
Public Property Let BackColor1(BackColor1 As OLE_COLOR): bCol1 = BackColor1: bNotOk = False: UserControl_Paint: End Property
Public Property Get BackColor2() As OLE_COLOR: BackColor2 = bCol2: End Property
Public Property Let BackColor2(BackColor2 As OLE_COLOR): bCol2 = BackColor2: bNotOk = False: UserControl_Paint: End Property
Public Property Get ShowBackGround() As Boolean: ShowBackGround = bShowBackGround: End Property
Public Property Let ShowBackGround(ShowBackGround As Boolean): bShowBackGround = ShowBackGround: bNotOk = False: UserControl_Paint: End Property
Public Property Get BackGradient() As GradientConstants: BackGradient = lBackGradient: End Property
Public Property Let BackGradient(BackGradient As GradientConstants): lBackGradient = BackGradient: bNotOk = False: UserControl_Paint: End Property
Public Property Get Enabled() As Boolean: Enabled = bEnable: End Property
Public Property Let Enabled(Enabled As Boolean)
bEnable = Enabled: bNotOk = False: UserControl_Paint: EnableControls
End Property
Public Property Get BorderColor() As OLE_COLOR: BorderColor = lBorderColor: End Property
Public Property Let BorderColor(BorderColor As OLE_COLOR): lBorderColor = BorderColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get DisplayBorder() As Boolean: DisplayBorder = bDisplayBorder: End Property
Public Property Let DisplayBorder(DisplayBorder As Boolean): bDisplayBorder = DisplayBorder: bNotOk = False: UserControl_Paint: End Property
Public Property Get BorderWidth() As Long: BorderWidth = lBWidth: End Property
Public Property Let BorderWidth(BorderWidth As Long): lBWidth = BorderWidth: bNotOk = False: UserControl_Paint: End Property
Public Property Get UnRefreshControl() As Boolean: UnRefreshControl = bUnRefreshControl: End Property
Public Property Let UnRefreshControl(UnRefreshControl As Boolean): bUnRefreshControl = UnRefreshControl: End Property
Public Property Get AreaHeight() As Long: AreaHeight = lAreaHeight: End Property
Public Property Get AreaWidth() As Long: AreaWidth = lAreaWidth: End Property
Public Property Get x() As Long: x = lX: End Property
Public Property Get y() As Long: y = lY: End Property
Public Property Let AreaHeight(AreaHeight As Long)
VS.Max = AreaHeight
lAreaHeight = AreaHeight: ReorganizeView
End Property
Public Property Let AreaWidth(AreaWidth As Long)
HS.Max = AreaWidth
lAreaWidth = AreaWidth: ReorganizeView
End Property
Public Property Let x(x As Long)
HS.Value = x
lX = x: ReorganizeView: Call MoveControls((OldValueH - x) / 2, 0): OldValueH = x
End Property
Public Property Let y(y As Long)
VS.Value = y
lY = y: ReorganizeView: Call MoveControls(0, (OldValueV - y) / 2): OldValueV = y
End Property
'//Propriétés sur les Scrolls
Public Property Get VScroll() As vkPrivateScroll
    Set VScroll = New vkPrivateScroll
    With VScroll
        .ArrowColor = VS.ArrowColor
        .BackColor = VS.BackColor
        .BorderColor = VS.BorderColor
        .DownColor = VS.DownColor
        .Enabled = VS.Enabled
        .EnableWheel = VS.EnableWheel
        .FrontColor = VS.FrontColor
        .hDC = VS.hDC
        .hWnd = VS.hWnd
        .LargeChange = VS.LargeChange
        .LargeChangeColor = VS.LargeChangeColor
        .Max = VS.Max
        .Min = VS.Min
        .MouseHoverColor = VS.MouseHoverColor
        .MouseInterval = VS.MouseInterval
        .ScrollHeight = VS.ScrollHeight
        .SmallChange = VS.SmallChange
        .UnRefreshControl = VS.UnRefreshControl
        .Value = VS.Value
        .WheelChange = VS.WheelChange
        .Width = VS.Width
        .Height = VS.Height
    End With
End Property
Public Property Let VScroll(VScroll As vkPrivateScroll)
    With VS
        .UnRefreshControl = True
        .ArrowColor = VScroll.ArrowColor
        .BackColor = VScroll.BackColor
        .BorderColor = VScroll.BorderColor
        .DownColor = VScroll.DownColor
        .Enabled = VScroll.Enabled
        .EnableWheel = VScroll.EnableWheel
        .FrontColor = VScroll.FrontColor
        .LargeChange = VScroll.LargeChange
        .LargeChangeColor = VScroll.LargeChangeColor
        .MouseHoverColor = VScroll.MouseHoverColor
        .MouseInterval = VScroll.MouseInterval
        .ScrollHeight = VScroll.ScrollHeight
        .SmallChange = VScroll.SmallChange
        .UnRefreshControl = VScroll.UnRefreshControl
        .Value = VScroll.Value
        .WheelChange = VScroll.WheelChange
        .UnRefreshControl = False
    End With
    Call VS.Refresh
    'bNotOk = False: Call UserControl_Paint
End Property
Public Property Get HScroll() As vkPrivateScroll
    Set HScroll = New vkPrivateScroll
    With HScroll
        .ArrowColor = HS.ArrowColor
        .BackColor = HS.BackColor
        .BorderColor = HS.BorderColor
        .DownColor = HS.DownColor
        .Enabled = HS.Enabled
        .EnableWheel = HS.EnableWheel
        .FrontColor = HS.FrontColor
        .hDC = HS.hDC
        .hWnd = HS.hWnd
        .LargeChange = HS.LargeChange
        .LargeChangeColor = HS.LargeChangeColor
        .Max = HS.Max
        .Min = HS.Min
        .MouseHoverColor = HS.MouseHoverColor
        .MouseInterval = HS.MouseInterval
        .ScrollHeight = HS.ScrollWidth
        .SmallChange = HS.SmallChange
        .UnRefreshControl = HS.UnRefreshControl
        .Value = HS.Value
        .WheelChange = HS.WheelChange
        .Width = HS.Width
        .Height = HS.Height
    End With
End Property
Public Property Let HScroll(HScroll As vkPrivateScroll)
    With HS
        .UnRefreshControl = True
        .ArrowColor = HScroll.ArrowColor
        .BackColor = HScroll.BackColor
        .BorderColor = HScroll.BorderColor
        .DownColor = HScroll.DownColor
        .Enabled = HScroll.Enabled
        .EnableWheel = HScroll.EnableWheel
        .FrontColor = HScroll.FrontColor
        .LargeChange = HScroll.LargeChange
        .LargeChangeColor = HScroll.LargeChangeColor
        .MouseHoverColor = HScroll.MouseHoverColor
        .MouseInterval = HScroll.MouseInterval
        .ScrollWidth = HScroll.ScrollHeight
        .SmallChange = HScroll.SmallChange
        .UnRefreshControl = HScroll.UnRefreshControl
        .Value = HScroll.Value
        .WheelChange = HScroll.WheelChange
        .UnRefreshControl = False
    End With
    Call HS.Refresh
   ' bNotOk = False: Call UserControl_Paint
End Property

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
'déplace tous les controles de X, Y
'=======================================================
Public Sub MoveControls(ByVal x As Long, ByVal y As Long)
Dim ctr As Control
    
    On Error Resume Next
    
    For Each ctr In UserControl.ContainedControls
        'traitement spécial pour les Lines
        If Not (TypeOf ctr Is VB.Line) Then
            With ctr
                .Left = .Left + x
                .Top = .Top + y
            End With
        Else
            'on change X1, Y1, X2 et Y2
            With ctr
                .X1 = .X1 + x
                .X2 = .X2 + x
                .Y1 = .Y1 + y
                .Y2 = .Y2 + y
            End With
        End If
    Next ctr
    
End Sub



'=======================================================
'PUBLIC SUB
'=======================================================

'=======================================================
'enable ou pas les controles qui sont dans le frame
'=======================================================
Private Sub EnableControls()
Dim ctr As Control
    
    On Error Resume Next
    
    For Each ctr In UserControl.ContainedControls
        ctr.Enabled = bEnable
    Next ctr
    
End Sub

'=======================================================
'on dessine tout
'=======================================================
Public Sub Refresh()
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
    Call OleTranslateColor(bCol1, 0, bCol1)
    Call OleTranslateColor(bCol2, 0, bCol2)
    
    
    If lBackStyle = TRANSPARENT Then
        'alors on créé une image dans le maskpicture
        
        'on dessine donc maintenant le bord
        If bDisplayBorder Then
                
            'on défini un brush
            hBrush = CreateSolidBrush(lBorderColor)
            
            'on définit une zone rectangulaire à bords arrondi
            hRgn = CreateRectRgn(0, 0, ScaleWidth / Screen.TwipsPerPixelX, _
                ScaleHeight / Screen.TwipsPerPixelY)
            
            'on dessine le contour
            Call FrameRgn(UserControl.hDC, hRgn, hBrush, lBWidth, lBWidth)

            'on détruit le brush et la zone
            Call DeleteObject(hBrush)
            Call DeleteObject(hRgn)
        End If
        
        UserControl.MaskPicture = UserControl.Image
    End If

    '//créé le gradient du reste du contrôle
    If lBackStyle = Opaque Then
        
        UserControl.BackStyle = 1
        
        If bShowBackGround Then
    
            'récupère les 3 composantes des deux couleurs
            
            If lBackGradient = None Then
                'pas de gradient
                'on dessine alors un rectangle
                Line (15, 15)-(ScaleWidth - 30, _
                    ScaleHeight - 30), bCol1, BF
            ElseIf lBackGradient = Horizontal Then
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
        End If
'    Else
'        'on récupère le maskpicture et on met le controle en transparent
'
'        UserControl.BackStyle = 0
'        UserControl.Picture = UserControl.MaskPicture
'
    End If
       
    
    '//on trace la bordure si opaque
    If lBackStyle = Opaque And bDisplayBorder Then
            
        'on défini un brush
        hBrush = CreateSolidBrush(lBorderColor)
        
        'on définit une zone rectangulaire à bords arrondi
        hRgn = CreateRectRgn(0, 0, ScaleWidth / Screen.TwipsPerPixelX, _
            ScaleHeight / Screen.TwipsPerPixelY)
        
        'on dessine le contour
        Call FrameRgn(UserControl.hDC, hRgn, hBrush, lBWidth, lBWidth)

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
'renvoie l'objet extender de ce usercontrol (pour les propertypages)
'=======================================================
Friend Property Get MyExtender() As Object
    Set MyExtender = UserControl.Extender
End Property
Friend Property Let MyExtender(MyExtender As Object)
    Set UserControl.Extender = MyExtender
End Property

Private Sub VS_Change(Value As Currency)
    
    RaiseEvent VScrollChange(Value)
    
    'on déplace tous les controles
    Call MoveControls(0, CLng(OldValueV - Value) / 2)
    
    OldValueV = Value
    
End Sub
Private Sub VS_Scroll()

    RaiseEvent VScrollScroll
    
    'on déplace tous les controles
    Call MoveControls(0, CLng(OldValueV - VS.Value) / 2)
    
    OldValueV = VS.Value
    
End Sub
Private Sub HS_Change(Value As Currency)

    RaiseEvent HScrollChange(Value)
    
    'on déplace tous les controles
    Call MoveControls(CLng(OldValueH - Value) / 2, 0)
    
    OldValueH = Value
    
End Sub
Private Sub HS_Scroll()

    RaiseEvent HScrollScroll
    
    'on déplace tous les controles
    Call MoveControls(CLng(OldValueH - HS.Value) / 2, 0)
    
    OldValueH = HS.Value

End Sub

'=======================================================
'on change la vue en fonction de X,Y et de la taille
'=======================================================
Public Sub ReorganizeView()

    '//on détermine si les scrolls doivent être visibles ou pas
    If Width > lAreaWidth Then
        'alors pas besoin de HScroll
        HS.Visible = False
    Else
        HS.Visible = True
    End If
    If Height > lAreaHeight Then
        'alors pas besoin de HScroll
        VS.Visible = False
    Else
        VS.Visible = True
    End If
    
    '//on replace les barres
    bNotResize = True
    Call UserControl_Resize
    bNotResize = False
    
    '//on place les barres au premier plan
'    Call HS.ZOrder(vbBringToFront)
'    Call VS.ZOrder(vbBringToFront)
End Sub
