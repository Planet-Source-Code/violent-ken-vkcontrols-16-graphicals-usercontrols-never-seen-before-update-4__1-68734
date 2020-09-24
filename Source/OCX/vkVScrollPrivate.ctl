VERSION 5.00
Begin VB.UserControl vkVScrollPrivate 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   600
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   0
      Top             =   120
   End
End
Attribute VB_Name = "vkVScrollPrivate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
Private mAsm(63) As Byte    'contient le code ASM
Private OldProc As Long     'adresse de l'ancienne window proc
Private objHwnd As Long     'handle de l'objet concerné
Private ET As TRACKMOUSEEVENTTYPE   'type pour le mouse_hover et le mouse_leave
Private IsMouseIn As Boolean    'si la souris est dans le controle

Private bCol As OLE_COLOR
Private lArrowColor As OLE_COLOR
Private lFrontColor As OLE_COLOR
Private lBorderColor As OLE_COLOR
Private lSelColor As OLE_COLOR
Private lLargeChangeColor As OLE_COLOR
Private lDownColor As OLE_COLOR
Private bEnable As Boolean
Private bNotOk As Boolean
Private bNotOk2 As Boolean
Private lScrollHeight As Byte
Private lMin As Currency
Private lMax As Currency
Private lValue As Currency
Private lSmallChange As Currency
Private lLargeChange As Currency
Private lWheelChange As Currency
Private bEnableWheel As Boolean
Private lPos1 As Long    'position haute du curseur
Private lPos2 As Long   'position basse du curseur
Private lGrise As Long
Private lUpMoused As Long
Private lDownMoused As Long
Private lMouseInterval As Long
Private lT As Long
Private lH As Long
Private nY As Long
Private n1 As Long
Private bUnRefreshControl As Boolean
Private bHasLeftOneTime As Boolean
Private bBlockVS As Boolean
Private bBlockValue As Boolean

'=======================================================
'EVENTS
'=======================================================
Public Event Change(Value As Currency)
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
Public Event Scroll()



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
                
                If lUpMoused = 1 Then
                    lUpMoused = 2: lValue = lValue - lSmallChange
                End If
                If lDownMoused = 1 Then
                    lDownMoused = 2: lValue = lValue + lSmallChange
                End If
                
                If bEnable Then
                    If y > 255 And y < lT Then
                        n1 = -1
                        lValue = lValue - lLargeChange
                    End If
                    If y > lT + lH And y < Height - 270 Then
                        n1 = 1
                        lValue = lValue + lLargeChange
                    End If
                End If
                
                Call ChangeValues
                RaiseEvent Change(lValue)
                
            RaiseEvent MouseDblClick(vbLeftButton, iShift, iControl, x, y)
        Case WM_LBUTTONDOWN
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                               
                If lUpMoused Then
                    lValue = lValue - lSmallChange
                    lUpMoused = 2: Timer1.Enabled = True: ChangeValues: RaiseEvent Change(lValue)
                End If
                If lDownMoused Then
                    lValue = lValue + lSmallChange
                    lDownMoused = 2: Timer1.Enabled = True: ChangeValues: RaiseEvent Change(lValue)
                End If
                
                If bEnable Then
                    If y > 255 And y < lT Then
                        lValue = lValue - lLargeChange
                        n1 = -1: ChangeValues: RaiseEvent Change(lValue)
                    ElseIf (y > lT + lH And y < Height - 270) Then
                        lValue = lValue + lLargeChange
                        n1 = 1: ChangeValues: RaiseEvent Change(lValue)
                    End If
                    If y > 255 And y < lT Then
                        Timer2.Enabled = True
                    End If
                    If y > lT + lH And y < Height - 270 Then
                        Timer2.Enabled = True
                    End If
                End If
                                
                RaiseEvent MouseDown(vbLeftButton, iShift, iControl, x, y)
        Case WM_LBUTTONUP
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
                n1 = 0
                If y > 255 And y < lT Then
                    Timer2.Enabled = False
                End If
                If y > lT + lH And y < Height - 270 Then
                    Timer2.Enabled = False
                End If
                
                Call Refresh
                If lUpMoused = 2 Then
                    lUpMoused = 1: Refresh: Timer1.Enabled = False
                End If
                If lDownMoused = 2 Then
                    lDownMoused = 1: Refresh: Timer1.Enabled = False
                End If
                
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
            lUpMoused = 0
            lDownMoused = 0
            If bHasLeftOneTime Then
                Call Refresh
            Else
                bHasLeftOneTime = True
            End If
        Case WM_MOUSEMOVE
            Call TrackMouseEvent(ET)
            
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
                If lUpMoused Then
                    If y > 255 Then
                        'on vire le cadre de sélection
                        lUpMoused = 0: Refresh
                    End If
                End If
                If lDownMoused Then
                    If y < Height - 270 Then
                        'on vire le cadre de sélection
                        lDownMoused = 0: Refresh
                    End If
                End If
                
                If lUpMoused = 0 And y <= 255 Then lUpMoused = 1: Refresh
                If lDownMoused = 0 And y >= Height - 270 Then lDownMoused = 1: Refresh
                
                If (wParam And MK_LBUTTON) = MK_LBUTTON Then z = vbLeftButton
                If (wParam And MK_RBUTTON) = MK_RBUTTON Then z = vbRightButton
                If (wParam And MK_MBUTTON) = MK_MBUTTON Then z = vbMiddleButton
                
                If z = vbLeftButton Then
                    'alors clic gauche enfoncé
                    If nY >= lT And nY <= lH + lT Then
                        'alors c'est sur le curseur O_o !
                        
                        RaiseEvent Scroll
                        
                        lT = lT + y - nY
                        
                        If lT <= 270 Then lT = 270
                        If lT >= Height - 285 - lH Then lT = Height - 285 - lH
                        
                        lValue = Int((lT - 255) * (lMax - lMin) / (Height - 510 - lH)) + lMin
                                          
                        If lT = 270 Then lValue = lMin
                        If lT = Height - 285 - lH Then lValue = lMax
                        
                        Call Refresh
                    End If
                   ' RaiseEvent Change(lValue)
                End If
                
                'sauvegarde la position
                nY = y
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
                If bEnableWheel Then Me.Value = Me.Value + lWheelChange
                RaiseEvent Change(lValue)
            Else
                RaiseEvent MouseWheel(WHEEL_UP)
                If bEnableWheel Then Me.Value = Me.Value - lWheelChange
                RaiseEvent Change(lValue)
            End If
        Case WM_PAINT
            bNotOk = True  'évite le clignotement lors du survol de la souris
    End Select
    
    'appel de la routine standard pour les autres messages
    WindowProc = CallWindowProc(OldProc, hWnd, uMsg, wParam, lParam)
    
End Function

Private Sub Timer1_Timer()
    
    If bEnable = False Then Exit Sub
    
    If lUpMoused = 2 Then
        'on clique sur le bouton haut
        lValue = lValue - lSmallChange
        Call ChangeValues
        RaiseEvent Change(lValue)
    ElseIf lDownMoused = 2 Then
        'bouton du bas
        lValue = lValue + lSmallChange
        Call ChangeValues
        RaiseEvent Change(lValue)
    End If
End Sub

Private Sub Timer2_Timer()
    
    If bEnable = False Or n1 = 0 Then Exit Sub
    
    If n1 = 1 Then
        'largechange en bas
        lValue = lValue + lLargeChange
        Call ChangeValues
        RaiseEvent Change(lValue)
    ElseIf n1 = -1 Then
        lValue = lValue - lLargeChange
        Call ChangeValues
        RaiseEvent Change(lValue)
    End If
        
    If lValue = lMax Or lValue = lMin Then Timer2.Enabled = False

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
        .ArrowColor = vbWhite
        .BackColor = vbWhite
        .BorderColor = &HFF8080
        .Enabled = True
        .EnableWheel = True
        .FrontColor = 15782079
        .LargeChange = 10
        .Max = 100
        .Min = 0
        .ScrollHeight = 10
        .SmallChange = 1
        .Value = 50
        .WheelChange = 3
        .DownColor = 12492429
        .MouseHoverColor = vbWhite
        .MouseInterval = 100
        .LargeChangeColor = 12492429
        .UnRefreshControl = True
    End With
    bNotOk2 = False
    Call UserControl_Paint  'refresh
    Timer1.Enabled = True
    Timer2.Enabled = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If bEnable = False Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyLeft, vbKeyUp
            lValue = lValue - SmallChange: ChangeValues: RaiseEvent Change(lValue)
        Case vbKeyRight, vbKeyDown
            lValue = lValue + SmallChange: ChangeValues: RaiseEvent Change(lValue)
        Case vbKeyPageUp
            lValue = lValue - LargeChange: ChangeValues: RaiseEvent Change(lValue)
        Case vbKeyPageDown
            lValue = lValue + LargeChange: ChangeValues: RaiseEvent Change(lValue)
        Case vbKeyEnd
            lValue = lMax: ChangeValues: RaiseEvent Change(lValue)
        Case vbKeyHome
            lValue = lMin: ChangeValues: RaiseEvent Change(lValue)
    End Select
        
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub UserControl_Terminate()
    'vire le subclassing
    If OldProc Then Call SetWindowLong(UserControl.hWnd, GWL_WNDPROC, OldProc)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("ArrowColor", Me.ArrowColor, vbWhite)
        Call .WriteProperty("BackColor", Me.BackColor, vbWhite)
        Call .WriteProperty("BorderColor", Me.BorderColor, &HFF8080)
        Call .WriteProperty("Enabled", Me.Enabled, True)
        Call .WriteProperty("EnableWheel", Me.EnableWheel, True)
        Call .WriteProperty("FrontColor", Me.FrontColor, 15782079)
        Call .WriteProperty("LargeChange", Me.LargeChange, 10)
        Call .WriteProperty("Max", Me.Max, 100)
        Call .WriteProperty("Min", Me.Min, 0)
        Call .WriteProperty("WheelChange", Me.WheelChange, 3)
        Call .WriteProperty("ScrollHeight", Me.ScrollHeight, 10)
        Call .WriteProperty("SmallChange", Me.SmallChange, 1)
        Call .WriteProperty("Value", Me.Value, 50)
        Call .WriteProperty("MouseHoverColor", Me.MouseHoverColor, vbWhite)
        Call .WriteProperty("DownColor", Me.DownColor, 12492429)
        Call .WriteProperty("MouseInterval", Me.MouseInterval, 100)
        Call .WriteProperty("LargeChangeColor", Me.LargeChangeColor, 12492429)
        Call .WriteProperty("UnRefreshControl", Me.UnRefreshControl, True)
    End With
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    bNotOk2 = True
    With PropBag
        Me.ArrowColor = .ReadProperty("ArrowColor", vbWhite)
        Me.BackColor = .ReadProperty("BackColor", vbWhite)
        Me.BorderColor = .ReadProperty("BorderColor", &HFF8080)
        Me.Enabled = .ReadProperty("Enabled", True)
        Me.EnableWheel = .ReadProperty("EnableWheel", True)
        Me.FrontColor = .ReadProperty("FrontColor", 15782079)
        Me.LargeChange = .ReadProperty("LargeChange", 10)
        Me.Max = .ReadProperty("Max", 100)
        Me.Min = .ReadProperty("Min", 0)
        Me.ScrollHeight = .ReadProperty("ScrollHeight", 10)
        Me.SmallChange = .ReadProperty("SmallChange", 1)
        Me.Value = .ReadProperty("Value", 50)
        Me.WheelChange = .ReadProperty("WheelChange", 3)
        Me.MouseHoverColor = .ReadProperty("MouseHoverColor", vbWhite)
        Me.DownColor = .ReadProperty("DownColor", 12492429)
        Me.MouseInterval = .ReadProperty("MouseInterval", 100)
        Me.LargeChangeColor = .ReadProperty("LargeChangeColor", 12492429)
        Me.UnRefreshControl = .ReadProperty("UnRefreshControl", True)
    End With
    bNotOk2 = False
    'Call UserControl_Paint  'refresh
    Call Refresh
    
    'le bon endroit pour lancer le subclassing
    'Call LaunchKeyMouseEvents
End Sub
Private Sub UserControl_Resize()
'    If Height < 800 Then Height = 800
  '  If Width < 255 Then Width = 255

    'lScrollHeight représente le pourcentage de la hauteur
    'calcule la hauteur du curseur
    lH = Int((Height - 510) * lScrollHeight / 100)

    Call ChangeValues
    'Call Refresh
End Sub

'=======================================================
'lance le subclassing
'=======================================================
Public Sub LaunchKeyMouseEvents()
                
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
Public Property Get BackColor() As OLE_COLOR: BackColor = bCol: End Property
Public Property Let BackColor(BackColor As OLE_COLOR): bCol = BackColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get BorderColor() As OLE_COLOR: BorderColor = lBorderColor: End Property
Public Property Let BorderColor(BorderColor As OLE_COLOR): lBorderColor = BorderColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get ArrowColor() As OLE_COLOR: ArrowColor = lArrowColor: End Property
Public Property Let ArrowColor(ArrowColor As OLE_COLOR): lArrowColor = ArrowColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get FrontColor() As OLE_COLOR: FrontColor = lFrontColor: End Property
Public Property Let FrontColor(FrontColor As OLE_COLOR): lFrontColor = FrontColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get Enabled() As Boolean: Enabled = bEnable: End Property
Public Property Let Enabled(Enabled As Boolean)
If bEnable <> Enabled Then
    bEnable = Enabled: bNotOk = False: UserControl_Paint
End If
End Property
Public Property Get EnableWheel() As Boolean: EnableWheel = bEnableWheel: End Property
Public Property Let EnableWheel(EnableWheel As Boolean): bEnableWheel = EnableWheel: bNotOk = False: UserControl_Paint: End Property
Public Property Get ScrollHeight() As Byte: ScrollHeight = lScrollHeight: End Property
Public Property Let ScrollHeight(ScrollHeight As Byte)
lScrollHeight = ScrollHeight
'lScrollHeight représente le pourcentage de la hauteur
'calcule la hauteur du curseur
lH = Int((Height - 510) * lScrollHeight / 100)
ChangeValues: bNotOk = False: UserControl_Paint
End Property
Public Property Get Min() As Currency: Min = lMin: End Property
Public Property Let Min(Min As Currency)
If lMin > lValue Then Exit Property
If lMin <> Min Then
    lMin = Min
    ChangeValues
    bNotOk = False: UserControl_Paint
End If
End Property
Public Property Get Max() As Currency: Max = lMax: End Property
Public Property Let Max(Max As Currency)
'If lMax < lValue Then Exit Property
If Max <> lMax Then
    lMax = Max
    ChangeValues
    bNotOk = False: UserControl_Paint
End If
End Property
Public Property Get Value() As Currency: Value = lValue: End Property
Public Property Let Value(Value As Currency)
If bBlockValue Then Exit Property
If Value <> lValue Then
    If bBlockVS = False Then RaiseEvent Change(lValue)
    lValue = Value: Call ChangeValues
End If
bBlockVS = False
End Property
Public Property Get SmallChange() As Currency: SmallChange = lSmallChange: End Property
Public Property Let SmallChange(SmallChange As Currency): lSmallChange = SmallChange: bNotOk = False: UserControl_Paint: End Property
Public Property Get LargeChange() As Currency: LargeChange = lLargeChange: End Property
Public Property Let LargeChange(LargeChange As Currency): lLargeChange = LargeChange: bNotOk = False: UserControl_Paint: End Property
Public Property Get WheelChange() As Currency: WheelChange = lWheelChange: End Property
Public Property Let WheelChange(WheelChange As Currency): lWheelChange = WheelChange: bNotOk = False: UserControl_Paint: End Property
Public Property Get DownColor() As OLE_COLOR: DownColor = lDownColor: End Property
Public Property Let DownColor(DownColor As OLE_COLOR): lDownColor = DownColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get MouseHoverColor() As OLE_COLOR: MouseHoverColor = lSelColor: End Property
Public Property Let MouseHoverColor(MouseHoverColor As OLE_COLOR): lSelColor = MouseHoverColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get MouseInterval() As Long: MouseInterval = lMouseInterval: End Property
Public Property Let MouseInterval(MouseInterval As Long): lMouseInterval = MouseInterval: Timer1.Interval = lMouseInterval: Timer2.Interval = lMouseInterval: End Property
Public Property Get LargeChangeColor() As OLE_COLOR: LargeChangeColor = lLargeChangeColor: End Property
Public Property Let LargeChangeColor(LargeChangeColor As OLE_COLOR): lLargeChangeColor = LargeChangeColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get UnRefreshControl() As Boolean: UnRefreshControl = bUnRefreshControl: End Property
Public Property Let UnRefreshControl(UnRefreshControl As Boolean): bUnRefreshControl = UnRefreshControl: End Property
Public Property Let BlockVS(BlockVS As Boolean): bBlockVS = BlockVS: End Property
Public Property Let BlockValue(BlockValue As Boolean): bBlockValue = BlockValue: End Property


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
'récupère la hauteur d'un caractère
'=======================================================
Private Function GetCharHeight() As Long
Dim Res As Long
    Res = GetTabbedTextExtent(UserControl.hDC, "A", 1, 0, 0)
    GetCharHeight = (Res And &HFFFF0000) \ &H10000
End Function

'=======================================================
'change la valeur Value
'=======================================================
Private Sub ChangeValues()
    
    If lValue > lMax Then lValue = lMax
    If lValue < lMin Then lValue = lMin
    
    'calcule le Top du curseur
    If lMax <> lMin Then _
        lT = By15(Int(Abs((Height - 510 - lH) * (lValue - lMin) / (lMax - lMin))) + 255)

    If lT <= 270 Then lT = 270
    If lT >= Height - 285 - lH Then lT = Height - 285 - lH
    
    'refresh le controle
    bNotOk = False: Call UserControl_Paint
End Sub

'=======================================================
'MAJ du controle
'=======================================================
Public Sub Refresh()
Dim x As Long
Dim y As Long

    If bUnRefreshControl Then Exit Sub
    
    '//on efface
    Call UserControl.Cls

    '//convertit les couleurs
    Call OleTranslateColor(lArrowColor, 0, lArrowColor)
    Call OleTranslateColor(lFrontColor, 0, lFrontColor)
    Call OleTranslateColor(bCol, 0, bCol)
    Call OleTranslateColor(lBorderColor, 0, lBorderColor)
    
    
    '//on trace les bords haut et bas, leur bordure et la bordure générale
    'contour général
    Line (0, 0)-(Width, Height), lBorderColor, BF
    'repeind intérieur en backcolor
    Line (15, 15)-(Width - 30, Height - 30), bCol, BF
    'zone haut
    Line (15, 15)-(Width - 30, 255), lFrontColor, BF
    'zone bas
    Line (15, Height - 285)-(Width - 30, Height - 30), lFrontColor, BF
    'lignes de séparation zones haut et bas
    Line (0, 270)-(Width, 270), lBorderColor
    Line (0, Height - 285)-(Width, Height - 285), lBorderColor
    

    
    '//trace les rectangles de sélection/pushed/rien
    Call DrawSelRectUp
    Call DrawSelRectDown
    
    
    '//on trace les flèches
    'si Enabled=false on met la couleur 10070188 aux flèches
    If bEnable Then
        UserControl.ForeColor = lArrowColor
    Else
        UserControl.ForeColor = 10070188
    End If
    x = (Width - 255) / 2 + 15
    'flèche du haut
    Line (105 + x, 90)-(120 + x, 90)
    Line (90 + x, 105)-(135 + x, 105)
    Line (75 + x, 120)-(150 + x, 120)
    Line (60 + x, 135)-(165 + x, 135)
    'en bas maintenant
    Line (105 + x, Height - 105)-(120 + x, Height - 105)
    Line (90 + x, Height - 120)-(135 + x, Height - 120)
    Line (75 + x, Height - 135)-(150 + x, Height - 135)
    Line (60 + x, Height - 150)-(165 + x, Height - 150)
    
    
    
    If bEnable = False Then
        'boulot terminé !
        Call UserControl.Refresh
        bNotOk = True
        Exit Sub
    End If
    
    
    
    '//on trace le curseur
    Line (15, lT)-(Width - 30, lT + lH), lBorderColor, BF  'bordure
    Line (15, lT + 15)-(Width - 30, lT + lH - 15), lFrontColor, BF
    
    
    
    '//on trace un rectangle de LargeChange
    If n1 = -1 Then
        'en haut
        Line (15, 285)-(Width - 30, lT - 15), lLargeChangeColor, BF
    ElseIf n1 = 1 Then
        'en bas
        Line (15, lT + lH + 15)-(Width - 30, Height - 300), lLargeChangeColor, BF
    End If
    

    '//on refresh le control
    Call UserControl.Refresh
    
    bNotOk = True
End Sub

'=======================================================
'trace le rectangle de sélection de la fleche du haut
'=======================================================
Private Sub DrawSelRectUp()

    'lUpMoused 1 (lignes blanches, survol) 2 (lignes foncées, pushed) 0 (rien)
    
    If bEnable = False Then Exit Sub

    Call OleTranslateColor(lSelColor, 0, lSelColor)
    Call OleTranslateColor(lDownColor, 0, lDownColor)
    
    Select Case lUpMoused
    
        Case 0
            Exit Sub
            
        Case 1
            'survol
            
            UserControl.ForeColor = lSelColor
            Line (15, 15)-(Width - 15, 15)
            Line (15, 15)-(15, 255)
            Line (15, 255)-(Width - 15, 255)
            Line (Width - 30, 240)-(Width - 30, 15)
            
        Case 2
            'clic
            
            Line (15, 15)-(Width - 30, 255), lDownColor, BF
            
    End Select

    Call UserControl.Refresh
End Sub

'=======================================================
'trace le rectangle de sélection de la fleche du bas
'=======================================================
Private Sub DrawSelRectDown()

    'lDownMoused 1 (lignes blanches, survol) 2 (lignes foncées, pushed) 0 (rien)
    
    If bEnable = False Then Exit Sub

    Call OleTranslateColor(lSelColor, 0, lSelColor)
    Call OleTranslateColor(lDownColor, 0, lDownColor)
    
    Select Case lDownMoused
    
        Case 0
            Exit Sub
            
        Case 1
            'survol
            
            UserControl.ForeColor = lSelColor
            Line (15, Height - 30)-(Width - 15, Height - 30)
            Line (15, Height - 30)-(15, Height - 270)
            Line (15, Height - 270)-(Width - 15, Height - 270)
            Line (Width - 30, Height - 270)-(Width - 30, Height - 30)
            
        Case 2
            'clic

            Line (15, Height - 270)-(Width - 30, Height - 30), lDownColor, BF

    End Select
    
End Sub

'=======================================================
'renvoie une valeur divisible par 15 (supérieure à l)
'=======================================================
Private Function By15(ByVal l As Currency) As Currency
    By15 = Int((l + 14) / 15) * 15
End Function

'=======================================================
'renvoie l'objet extender de ce usercontrol (pour les propertypages)
'=======================================================
Friend Property Get MyExtender() As Object
    Set MyExtender = UserControl.Extender
End Property
Friend Property Let MyExtender(MyExtender As Object)
    Set UserControl.Extender = MyExtender
End Property


