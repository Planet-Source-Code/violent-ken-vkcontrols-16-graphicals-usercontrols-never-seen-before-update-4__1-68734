VERSION 5.00
Begin VB.UserControl vkBar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   PropertyPages   =   "vkBar.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "vkBar.ctx":0049
   Begin VB.PictureBox frontImg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3360
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox imgNull 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2160
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   3
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox tpn 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   600
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   185
      TabIndex        =   2
      Top             =   1800
      Width           =   2775
   End
   Begin VB.PictureBox pct 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   480
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   185
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.PictureBox backImg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1200
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "vkBar"
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

Private lBackColorTop As Long 'couleur dégradé 1
Private lBackColorBottom As Long 'couleur dégradé 2
Private lLeftColor As Long 'couleur valeur dégradé 1
Private lRightColor As Long 'couleur valeur dégradé 2
Private dMin As Double 'minimum
Private dMax As Double 'maximum
Private bIsInteractive As Boolean 'contrôle intéractif ou non
Private bShowLabel As Label_Mode 'type d'affichae
Private dValue As Double 'value
Private mdDeg As Mode_Degrade   'type de dégradé
Private btButton As Button_Type 'bouton d'interaction
Private brndColor As Boolean    'couleur du contour
Private lLabelColor As Long 'forecolor
Private lPercentDecimal As Long 'nombre de décimales au pourcentage (label)
Private b3D As Border  'affichage en 3D
Private taAlign As Text_Alignment   'alignement du texte
Private lOSx As Long    'offsetX
Private lOSy As Long    'offsetY
Private bNotOk As Boolean
Private bNotOk2 As Boolean
Private bUnRefreshControl As Boolean


'=======================================================
'EVENTS publics
'=======================================================
Public Event Change(NewValue As Double, OldValue As Double)
Attribute Change.VB_Description = "Happens when Value is changed"
Public Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Happens when a key is pressed"
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Happens when a key is down"
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Happens when a key is up"
Public Event Click()
Attribute Click.VB_Description = "Happens when control gets a click (left button)"
Public Event DblClick()
Attribute DblClick.VB_Description = "Happens when control gets a dblclick (left button)"
Public Event InteractionComplete(NewValue As Double, OldValue As Double)
Attribute InteractionComplete.VB_Description = "Happens when interaction is completed."
Public Event ValueIsMax(Value As Double)
Attribute ValueIsMax.VB_Description = "Happens when Value=Max"
Public Event ValueIsMin(Value As Double)
Attribute ValueIsMin.VB_Description = "Happens when Value=Min"
Public Event MouseWheel(WheelSens As Wheel_Sens)
Attribute MouseWheel.VB_Description = "Happens when control gets a wheel"
Public Event MouseHover()
Attribute MouseHover.VB_Description = "Happens when mouse enters control"
Public Event MouseLeave()
Attribute MouseLeave.VB_Description = "Happens when mouse leaves control"
Public Event MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
Attribute MouseDown.VB_Description = "Happens when control getsa click"
Public Event MouseUp(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
Attribute MouseUp.VB_Description = "Happens when control gets a mouseup"
Public Event MouseDblClick(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
Attribute MouseDblClick.VB_Description = "Happens when control gets a dblclick"
Public Event MouseMove(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
Attribute MouseMove.VB_Description = "Happens when mouse moves on control"




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
Attribute WindowProc.VB_Description = "Internal proc for subclassing"
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

Private Sub UserControl_InitProperties()

    'valeurs par défaut
    bNotOk2 = True
    With Me
        .Min = 1
        .Max = 100
        .Value = 50
        .InteractiveControl = False
        .DisplayLabel = PercentageMode
        .RightColor = &HE4C6B5
        .LeftColor = &HC85A21
        .BackColorBottom = &HFBFBFB
        .BackColorTop = &HDCDCDC
        .GradientMode = AllLengh
        .InteractiveButton = LeftButton
        .DisplayBorder = True
        .BorderColor = &HFF8080
        .ForeColor = &H404040
        .Decimals = 2
        .BorderStyle = 0
        Set .Font = Ambient.Font
        .Alignment = MiddleCenter
        .OffSetX = 0
        .UnRefreshControl = False
        .OffSetY = 0
    End With
    bNotOk2 = False
    
    'refresh value
    Call Refresh
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    bNotOk2 = True
    With PropBag
        Set Me.Font = .ReadProperty("Font", Ambient.Font)
        Me.DisplayBorder = .ReadProperty("DisplayBorder", True)
        Me.Decimals = .ReadProperty("Decimals", 2)
        Me.BorderColor = .ReadProperty("BorderColor", &HFF8080)
        Me.BackColorTop = .ReadProperty("BackColorTop", &HDCDCDC)
        Me.BackColorBottom = .ReadProperty("BackColorBottom", &HFBFBFB)
        Me.LeftColor = .ReadProperty("LeftColor", &HC85A21)
        Me.RightColor = .ReadProperty("RightColor", &HE4C6B5)
        Me.Min = .ReadProperty("Min", 1)
        Me.Max = .ReadProperty("Max", 100)
        Me.OffSetX = .ReadProperty("OffSetX", 0)
        Me.OffSetY = .ReadProperty("OffSetY", 0)
        Me.Value = .ReadProperty("Value", 50)
        Me.InteractiveControl = .ReadProperty("InteractiveControl", False)
        Me.DisplayLabel = .ReadProperty("DisplayLabel", PercentageMode)
        Me.GradientMode = .ReadProperty("GradientMode", AllLengh)
        Me.ForeColor = .ReadProperty("ForeColor", &H404040)
        Me.BorderStyle = .ReadProperty("BorderStyle", 0)
        Me.Alignment = .ReadProperty("Alignment", 5)
        Me.InteractiveButton = .ReadProperty("InteractiveButton", LeftButton)
        Set backImg.Picture = .ReadProperty("BackPicture", imgNull.Picture)
        Set frontImg.Picture = .ReadProperty("FrontPicture", imgNull.Picture)
        Me.UnRefreshControl = .ReadProperty("UnRefreshControl", False)
    End With
    If backImg.Picture <> 0 Then Set tpn.Picture = backImg.Picture
    bNotOk2 = False
    'Call UserControl_Paint  'refresh
    Call Refresh
        
    'le bon endroit pour lancer le subclassing
    Call LaunchKeyMouseEvents
End Sub

Private Sub UserControl_Terminate()
    'vire le subclassing
    If OldProc Then Call SetWindowLong(UserControl.hWnd, GWL_WNDPROC, OldProc)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("BorderColor", Me.BorderColor, &HFF8080)
        Call .WriteProperty("DisplayBorder", Me.DisplayBorder, True)
        Call .WriteProperty("Decimals", Me.Decimals, 2)
        Call .WriteProperty("BackColorTop", Me.BackColorTop, &HDCDCDC)
        Call .WriteProperty("BackColorBottom", Me.BackColorBottom, &HFBFBFB)
        Call .WriteProperty("LeftColor", Me.LeftColor, &HC85A21)
        Call .WriteProperty("RightColor", Me.RightColor, &HE4C6B5)
        Call .WriteProperty("Min", Me.Min, 1)
        Call .WriteProperty("Max", Me.Max, 100)
        Call .WriteProperty("OffSetX", Me.OffSetX, 0)
        Call .WriteProperty("OffSetY", Me.OffSetY, 0)
        Call .WriteProperty("Alignment", Me.Alignment, 5)
        Call .WriteProperty("Value", Me.Value, 50)
        Call .WriteProperty("InteractiveControl", Me.InteractiveControl, False)
        Call .WriteProperty("DisplayLabel", Me.DisplayLabel, PercentageMode)
        Call .WriteProperty("GradientMode", Me.GradientMode, AllLengh)
        Call .WriteProperty("UnRefreshControl", Me.UnRefreshControl, False)
        Call .WriteProperty("ForeColor", Me.ForeColor, &H404040)
        Call .WriteProperty("BorderStyle", Me.BorderStyle, 0)
        Call .WriteProperty("InteractiveButton", Me.InteractiveButton, LeftButton)
        Call .WriteProperty("BackPicture", UserControl.backImg, imgNull.Picture)
        Call .WriteProperty("FrontPicture", UserControl.frontImg, imgNull.Picture)
        Call .WriteProperty("Font", Me.Font, Ambient.Font)
    End With
End Sub

'=======================================================
'PROPERTIES
'=======================================================
Public Property Get Alignment() As Text_Alignment: Alignment = taAlign: End Property
Attribute Alignment.VB_Description = "Text alignment"
Public Property Let Alignment(Alignment As Text_Alignment): taAlign = Alignment: bNotOk = False: Refresh: End Property
Public Property Get ForeColor() As OLE_COLOR: ForeColor = lLabelColor: End Property
Attribute ForeColor.VB_Description = "Text color"
Public Property Let ForeColor(ForeColor As OLE_COLOR): lLabelColor = ForeColor: Refresh: End Property
Public Property Get Value() As Double: Value = dValue: End Property
Attribute Value.VB_Description = "Current value"
Attribute Value.VB_MemberFlags = "200"
Public Property Let Value(Value As Double): Dim lOld As Double
    lOld = dValue
    If Value < Min Then
        dValue = Min
    ElseIf Value > dMax Then
        dValue = dMax
    Else
        dValue = Value
    End If
    RaiseEvent Change(dValue, lOld)
    If dValue = dMin Then RaiseEvent ValueIsMin(dMin)
    If dValue = dMax Then RaiseEvent ValueIsMax(dMax)
    bNotOk = False: Refresh
End Property
Public Property Get InteractiveButton() As Button_Type: InteractiveButton = btButton: End Property
Attribute InteractiveButton.VB_Description = "Define wich button is used to Interactive option"
Public Property Let InteractiveButton(InteractiveButton As Button_Type): btButton = InteractiveButton: End Property
Public Property Get InteractiveControl() As Boolean: InteractiveControl = bIsInteractive: End Property
Attribute InteractiveControl.VB_Description = "Active or not the Interactive option (allows to change value by using mouse)"
Public Property Let InteractiveControl(InteractiveControl As Boolean): bIsInteractive = InteractiveControl: End Property
Public Property Get DisplayLabel() As Label_Mode: DisplayLabel = bShowLabel: End Property
Attribute DisplayLabel.VB_Description = "Define how to display label"
Public Property Let DisplayLabel(DisplayLabel As Label_Mode): bShowLabel = DisplayLabel: bNotOk = False: Refresh:: End Property
Public Property Get Font() As StdFont: Set Font = UserControl.Font: End Property
Attribute Font.VB_Description = "Text font"
Public Property Set Font(Font As StdFont): Set UserControl.Font = Font: bNotOk = False: Refresh: End Property
Public Property Get DisplayBorder() As Boolean: DisplayBorder = brndColor: UserControl_Resize: End Property
Attribute DisplayBorder.VB_Description = "Display or not the border"
Public Property Let DisplayBorder(DisplayBorder As Boolean): brndColor = DisplayBorder: bNotOk = False: Refresh: End Property
Public Property Get BackPicture() As StdPicture
Attribute BackPicture.VB_Description = "Define the picture wich is displayed on the back of the control"
'BackPicture = lLeftColor
    Set BackPicture = tpn.Picture
    Set BackPicture = backImg.Picture
End Property
Public Property Set BackPicture(ByVal BackPicture As StdPicture)
    Set tpn.Picture = BackPicture
    Set backImg.Picture = BackPicture
End Property
Public Property Get FrontPicture() As StdPicture
Attribute FrontPicture.VB_Description = "Define the picture wich is displayed on the top of the control"
    Set FrontPicture = frontImg.Picture
End Property
Public Property Set FrontPicture(ByVal FrontPicture As StdPicture): Set frontImg.Picture = FrontPicture: End Property
Public Property Get LeftColor() As OLE_COLOR: LeftColor = lLeftColor: End Property
Attribute LeftColor.VB_Description = "Left color of the gradient"
Public Property Let LeftColor(LeftColor As OLE_COLOR): lLeftColor = LeftColor: bNotOk = False: Refresh: End Property
Public Property Get BorderStyle() As Border: BorderStyle = b3D: End Property
Attribute BorderStyle.VB_Description = "3D effect or not"
Public Property Let BorderStyle(BorderStyle As Border): b3D = BorderStyle: pct.BorderStyle = b3D: End Property
Public Property Get Decimals() As Long: Decimals = lPercentDecimal: End Property
Attribute Decimals.VB_Description = "Number of decimal to use to display percentage"
Public Property Let Decimals(Decimals As Long)
    If Decimals >= 23 Then
        lPercentDecimal = 22
    Else
        lPercentDecimal = Decimals
    End If
    bNotOk = False: Refresh
End Property
Public Property Get RightColor() As OLE_COLOR: RightColor = lRightColor: End Property
Attribute RightColor.VB_Description = "Right color of the gradient"
Public Property Let RightColor(RightColor As OLE_COLOR): lRightColor = RightColor: bNotOk = False: Refresh: End Property
Public Property Get BackColorTop() As OLE_COLOR: BackColorTop = lBackColorTop: End Property
Attribute BackColorTop.VB_Description = "Top backcolor"
Public Property Let BackColorTop(BackColorTop As OLE_COLOR): lBackColorTop = BackColorTop: bNotOk = False: Refresh: End Property
Public Property Get BackColorBottom() As OLE_COLOR: BackColorBottom = lBackColorBottom: End Property
Attribute BackColorBottom.VB_Description = "Bottom backcolor"
Public Property Let BackColorBottom(BackColorBottom As OLE_COLOR): lBackColorBottom = BackColorBottom: bNotOk = False: Refresh: End Property
Public Property Get OffSetX() As Long: OffSetX = lOSx: End Property
Attribute OffSetX.VB_Description = "Offset (pixels) of text"
Public Property Let OffSetX(OffSetX As Long): lOSx = OffSetX: bNotOk = False: Refresh: End Property
Public Property Get OffSetY() As Long: OffSetY = lOSy: End Property
Attribute OffSetY.VB_Description = "Offset (pixels) of text"
Public Property Let OffSetY(OffSetY As Long): lOSy = OffSetY: bNotOk = False: Refresh: End Property
Public Property Get Min() As Double: Min = dMin: End Property
Attribute Min.VB_Description = "Min of the range"
Public Property Let Min(Min As Double): dMin = Min: End Property
Public Property Get Max() As Double: Max = dMax: End Property
Attribute Max.VB_Description = "Max of the range"
Public Property Let Max(Max As Double)
If Max > dMin Then dMax = Max
End Property
Public Property Get BorderColor() As OLE_COLOR: BorderColor = UserControl.BackColor: End Property
Attribute BorderColor.VB_Description = "Color of the border"
Public Property Let BorderColor(BorderColor As OLE_COLOR): UserControl.BackColor = BorderColor: bNotOk = False: Refresh: End Property
Public Property Get GradientMode() As Mode_Degrade: GradientMode = mdDeg: End Property
Attribute GradientMode.VB_Description = "Define how to display gradient"
Public Property Let GradientMode(GradientMode As Mode_Degrade): mdDeg = GradientMode: bNotOk = False: Refresh: End Property
Public Property Get hWnd() As Long: hWnd = UserControl.hWnd: End Property
Attribute hWnd.VB_Description = "Handle of the control"
Public Property Get UnRefreshControl() As Boolean: UnRefreshControl = bUnRefreshControl: End Property
Attribute UnRefreshControl.VB_Description = "Prevent to refresh control"
Public Property Let UnRefreshControl(UnRefreshControl As Boolean): bUnRefreshControl = UnRefreshControl: End Property


'=======================================================
'EVENEMENTS SIMPLES
'=======================================================
Private Sub pct_Click(): RaiseEvent Click: End Sub
Private Sub pct_DblClick(): RaiseEvent DblClick: End Sub
Private Sub pct_KeyDown(KeyCode As Integer, Shift As Integer): RaiseEvent KeyDown(KeyCode, Shift): End Sub
Private Sub pct_KeyPress(KeyAscii As Integer): RaiseEvent KeyPress(KeyAscii): End Sub
Private Sub pct_KeyUp(KeyCode As Integer, Shift As Integer): RaiseEvent KeyUp(KeyCode, Shift): End Sub



Private Sub pct_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'actionne l'interaction
Dim lOld As Double

    'pour ne pas sortir du composant
    With pct
        If x > .ScaleWidth Then x = .ScaleWidth
        If x < 0 Then x = 0
        If y > .ScaleHeight Then y = .ScaleHeight
        If y < 0 Then y = 0
    End With
    
    If bIsInteractive And btButton = Button Then
        'alors on change la valeur
        
        lOld = dValue
        dValue = (dMax / pct.ScaleWidth) * x
        
        'met à 100% quand on sélectionne tout à droite
        If x = (pct.ScaleWidth - 1) Then dValue = dMax
        
        bNotOk = False
        Call Refresh 'réaffiche la barre

        'évênements
        RaiseEvent InteractionComplete(dValue, lOld)
        RaiseEvent Change(dValue, lOld)
        If dValue = dMin Then RaiseEvent ValueIsMin(dMin)
        If dValue = dMax Then RaiseEvent ValueIsMax(dMax)
        
    End If
    
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

Private Sub UserControl_Resize()
'resize le composant
Dim lDif As Long    'marge

    'calcule la marge (=0 si pas de bordure)
    If brndColor Then lDif = 30 Else lDif = 0

    'redimensionne les composants
    With pct
        .Left = lDif / 2
        .Top = lDif / 2
        .Height = Height - lDif
        .Width = Width - lDif
    End With
    With tpn
        .Left = -75000
        .Height = Height - lDif
        .Width = Width - lDif
    End With
    With backImg
        .Left = pct.Left
        .Top = pct.Top
        .Width = pct.Width
        .Height = pct.Height
    End With

    bNotOk = False
    Call Refresh  'refresh
End Sub

'=======================================================
'créé le dégradé
'=======================================================
Private Sub Degrader()
Dim pxlHeight As Long   'hauteur (pixel) du picturebox
Dim pxlWidth As Long    'largeur (pixel) du picturebox
Dim x As Long
Dim y As Long
Dim rUp As Long 'composante rouge couleur du haut
Dim gUp As Long 'composante verte couleur du haut
Dim bUp As Long 'composante bleue couleur du haut
Dim rDown As Long 'composante rouge couleur du bas
Dim gDown As Long 'composante verte couleur du bas
Dim bDown As Long 'composante bleue couleur du bas
Dim dIncrRed As Double  'incrémentation de la composante rouge
Dim dIncrGreen As Double  'incrémentation de la composante verte
Dim dIncrBlue As Double  'incrémentation de la composante bleue


    'on trace dans un picturebox tampon (évite de redissiner ce dégradé à chaque refresh)
    If Me.BackPicture <> 0 Then
        'alors une picture est affichée en fond, donc on redessine pas
        Exit Sub
    End If
    
    'clear Picture
    Call tpn.Cls
    
    'récupère les valeurs RBG des deux couleurs du dégradé
    Call LongToRGB(BackColorTop, rUp, gUp, bUp)
    Call LongToRGB(BackColorBottom, rDown, gDown, bDown)
    
    'dans le cas où ce sont des couleurs système
    Call OleTranslateColor(BackColorTop, 0, BackColorTop)
    Call OleTranslateColor(BackColorBottom, 0, BackColorBottom)
    
    'récupère les dimensions (pixel) du picturebox tampon
    pxlHeight = pct.ScaleHeight
    pxlWidth = pct.ScaleWidth
    
    'calcule les incrémentations pour chaque composante
    dIncrRed = (rUp - rDown) / pxlHeight
    dIncrGreen = (gUp - gDown) / pxlHeight
    dIncrBlue = (bUp - bDown) / pxlHeight
    
    'trace le dégradé
    For x = 0 To pxlHeight
        tpn.ForeColor = RGB(CByte(rDown + x * dIncrRed), CByte(gDown + x * dIncrGreen), CByte(bDown + x * dIncrBlue))
        tpn.Line (0, x)-(pxlWidth, x)
    Next x
    
End Sub

'=======================================================
'rafraichit le contrôle
'=======================================================
Private Sub Refresh()
Dim pxlHeight As Long   'hauteur du picturebox
Dim pxlWidth As Long    'largeur du picturebox
Dim x As Long
Dim y As Long
Dim lValueWidth As Double   'largeur (pixel) de la valeur
Dim lPlage As Double  'nombre de valeurs différentes possibles pour value (en long)
Dim rLeft As Long 'composante rouge couleur de gauche
Dim gLeft As Long 'composante verte couleur de gauche
Dim bLeft As Long 'composante bleue couleur de gauche
Dim rRight As Long 'composante rouge couleur de droite
Dim gRight As Long 'composante verte couleur de droite
Dim bRight As Long 'composante bleue couleur de droite
Dim dIncrRed As Double  'incrémentation de la composante rouge
Dim dIncrGreen As Double  'incrémentation de la composante verte
Dim dIncrBlue As Double  'incrémentation de la composante bleue
Dim lDif As Long    'marge de resizement
Dim txtWidth As Long    'largeur du texte
Dim txtHeight As Long    'hauteur du texte
Dim sText As String 'texte à afficher
Dim lRet As Long    'retour de l'API

    
    'On Error Resume Next
    
    If bNotOk Or bNotOk2 Then Exit Sub
    
    '//affichage du dégradé
    Call Degrader
    
    'obtient les dimensions du picturebox
    With pct
        pxlHeight = pct.ScaleHeight
        pxlWidth = pct.ScaleWidth
    End With
    
    If frontImg.Picture <> 0 Then
        'alors on a une image à afficher dans la barre de progression
        
        'calcule la largeur (pixel) de la valeur à afficher
        lPlage = dMax - dMin
        
        If lPlage = 0 Then Exit Sub 'pas encore initialisé
    
        'largeur de la picture à poser
        lValueWidth = ((dValue - dMin) / lPlage) * pxlWidth + 0.0001
                
        'efface la picturebox
        Call pct.Cls
        'ajoute le dégradé de fond
        pct.Picture = tpn.Image
        
        'plaque la picturebox de devant sur la picturebox contenant la barre
        Call StretchBlt(pct.hDC, 0, 0, Int(Screen.TwipsPerPixelX * lValueWidth), pct.Height, frontImg.hDC, 0, _
        0, frontImg.Width, frontImg.Height, &HCC0020)
        
        'pct.Picture = frontImg.Picture
        
        GoTo BarreDone
    End If
        
    'efface la picturebox
    Call pct.Cls
    
    'ajoute le dégradé de fond
    pct.Picture = tpn.Image
    
    'calcule la largeur (pixel) de la valeur à afficher
    lPlage = dMax - dMin
    
    If lPlage = 0 Then Exit Sub 'pas encore initialisé
    
    lValueWidth = ((dValue - dMin) / lPlage) * pxlWidth + 0.0001
    
    'récupère les valeurs RBG des deux couleurs du dégradé
    Call LongToRGB(LeftColor, rLeft, gLeft, bLeft)
    Call LongToRGB(RightColor, rRight, gRight, bRight)
    
    'calcule le pas en fonction du mode de dégradé
    If mdDeg = AllLengh Then
        'dégradé sur toute la longueur
        'calcul des incrémentations
        dIncrRed = (rLeft - rRight) / pxlWidth
        dIncrGreen = (gLeft - gRight) / pxlWidth
        dIncrBlue = (bLeft - bRight) / pxlWidth
    Else
        'dégradé uniquement sur la plage affichée
        dIncrRed = (rLeft - rRight) / lValueWidth
        dIncrGreen = (gLeft - gRight) / lValueWidth
        dIncrBlue = (bLeft - bRight) / lValueWidth
    End If
    
    'affichage des dégradés
    For x = 0 To Int(lValueWidth)
        pct.ForeColor = RGB(CByte(rLeft - x * dIncrRed), CByte(gLeft - x * dIncrGreen), CByte(bLeft - x * dIncrBlue))
        pct.Line (x, 0)-(x, pxlHeight)
    Next x
    
    
BarreDone:
    
    '//affichage (ou pas) du texte
    
    'affiche le texte dans le label (en autosize) pour pouvoir _
    'calculer la largeur à prévoir
    'If bShowLabel = No Then GoTo NoText
    If bShowLabel = PercentageMode Then
        'pourcentage
        lValueWidth = Round(100 * (dValue - dMin) / lPlage, lPercentDecimal)
        If lValueWidth < 0 Then lValueWidth = 0
        sText = CStr(lValueWidth) & " %"
    ElseIf bShowLabel = ValueMode Then
        'valeur
        sText = CStr(Round(dValue, lPercentDecimal))
    ElseIf bShowLabel = Steps Then
        'avancement en pas
        sText = CStr(Round(dValue, lPercentDecimal)) & "/" & CStr(dMax)
    End If

    
    'récupère la dimension (en pixels)
    txtWidth = TextWidth(sText) / Screen.TwipsPerPixelX
    txtHeight = TextHeight(sText) / Screen.TwipsPerPixelY
    
    'affiche le texte dans pct en le centrant
    
    'affichage du texte
    pct.ForeColor = ForeColor
    Set pct.Font = UserControl.Font
    
    'positionnement
    With pct
        Select Case taAlign
            Case MiddleCenter
                .CurrentX = Int((pxlWidth - txtWidth) / 2) + lOSx
                .CurrentY = Int((pxlHeight - txtHeight) / 2) + lOSy
            Case MiddleLeft
                .CurrentX = lOSx
                .CurrentY = Int((pxlHeight - txtHeight) / 2) + lOSy
            Case MiddleRight
                .CurrentX = pxlWidth - txtWidth + lOSx
                .CurrentY = Int((pxlHeight - txtHeight) / 2) + lOSy
            Case TopLeft
                .CurrentX = lOSx
                .CurrentY = lOSy
            Case TopCenter
                .CurrentX = Int((pxlWidth - txtWidth) / 2) + lOSx
                .CurrentY = lOSy
            Case TopRight
                .CurrentX = pxlWidth - txtWidth + lOSx
                .CurrentY = lOSy
            Case BottomLeft
                .CurrentX = lOSx
                .CurrentY = pxlHeight - txtHeight + lOSy
            Case BottomRight
                .CurrentX = pxlWidth - txtWidth + lOSx
                .CurrentY = pxlHeight - txtHeight + lOSy
            Case BottomCenter
                .CurrentX = Int((pxlWidth - txtWidth) / 2) + lOSx
                .CurrentY = pxlHeight - txtHeight + lOSy
        End Select


        pct.Print sText
    End With
    
    
NoText:
    
    '//resize le composant
    
    'calcule la marge (=0 si pas de bordure)
    If brndColor Then lDif = 30 Else lDif = 0

    'redimensionne les composants
    With pct
        .Left = lDif / 2
        .Top = lDif / 2
        .Height = Height - lDif
        .Width = Width - lDif
    End With
    With tpn
        .Left = -75000
        .Height = Height - lDif
        .Width = Width - lDif
    End With

    bNotOk = True
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
'renvoie l'objet extender de ce usercontrol (pour les propertypages)
'=======================================================
Friend Property Get MyExtender() As Object
    Set MyExtender = UserControl.Extender
End Property
Friend Property Let MyExtender(MyExtender As Object)
    Set UserControl.Extender = MyExtender
End Property
