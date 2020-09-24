VERSION 5.00
Begin VB.UserControl vkCheck 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   PropertyPages   =   "vkCheckBox.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "vkCheckBox.ctx":0039
   Begin VB.Image Image1 
      Height          =   195
      Left            =   480
      Picture         =   "vkCheckBox.ctx":034B
      Top             =   600
      Visible         =   0   'False
      Width           =   1170
   End
End
Attribute VB_Name = "vkCheck"
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
Private lForeColor As OLE_COLOR
Private bCol As OLE_COLOR
Private bEnable As Boolean
Private tVal As CheckBoxConstants
Private sCaption As String
Private bNotOk As Boolean
Private bHasFocus As Boolean
Private bNotOk2 As Boolean
Private tAlig As AlignmentConstants
Private bUnRefreshControl As Boolean
Private bHasLeftOneTime As Boolean
Private bUnicode As Boolean


'=======================================================
'EVENTS
'=======================================================
Public Event Click()
Attribute Click.VB_Description = "Happens when control gets a click (left button)"
Public Event Change(Value As CheckBoxConstants)
Attribute Change.VB_Description = "Happens when value changes"
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Happens when a key is down"
Public Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Happens when a key is pressed"
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Happens when a key is up"
Public Event MouseHover()
Attribute MouseHover.VB_Description = "Happens when mouse enters control"
Public Event MouseLeave()
Attribute MouseLeave.VB_Description = "Happens when mouse leaves control"
Public Event MouseWheel(Sens As Wheel_Sens)
Attribute MouseWheel.VB_Description = "Happens when control gets a wheel"
Public Event MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
Attribute MouseDown.VB_Description = "Happens when control gets a click"
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
                
                Call ChangeValue
            RaiseEvent MouseDblClick(vbLeftButton, iShift, iControl, x, y)
        Case WM_LBUTTONDOWN
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                Call ChangeValue
                
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
                Call Refresh
            End If
        Case WM_MOUSELEAVE
            RaiseEvent MouseLeave
            IsMouseIn = False
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

'=======================================================
'défini les touches de raccourci (avec '&')
'=======================================================
Private Sub SetAccessKeys()
Dim a As Long

    'récupère la position du '&' en partant de la fin
    a = InStrRev(sCaption, "&")
    
    'si le '&' existe et n'est pas tout à la fin
    If a <> Len(sCaption) And a <> 0 Then
        'on récupère le caractère qui est juste après
        AccessKeys = Mid$(sCaption, a + 1, 1)
    End If
End Sub

'=======================================================
'Change la value du control
'=======================================================
Private Sub ChangeValue()
    
    If bEnable = False Then Exit Sub
        
    If tVal = vbChecked Then
        tVal = vbUnchecked
    ElseIf tVal = vbUnchecked Then
        tVal = vbChecked
    End If
    
    RaiseEvent Change(tVal)
    RaiseEvent Click
    
    bNotOk = False: Call UserControl_Paint
    
End Sub

Private Sub UserControl_GotFocus()
'alors on va tracer un BÔ rectangle de sélection
Dim R As RECT
Dim y As Long
Dim x As Long
    
     If bEnable = False Then
        'on ne garde pas le focus
        Call SendKeys("{Tab}")
        Exit Sub
    End If
    
    'on a alors le focus
    bHasFocus = True
    
    '//on dessine le rectangle de focus
    'une zone rectangulaire
    y = (ScaleHeight / Screen.TwipsPerPixelY - GetCharHeight) / 2
    Call SetRect(R, 17, y - 1, TextWidth(sCaption) / Screen.TwipsPerPixelX + 23, y + _
        GetCharHeight + 2)
    'dessine
    Call DrawFocusRect(UserControl.hDC, R)
    
    If lBackStyle = [TRANSPARENT] Then
        'transparent
        With UserControl
            .BackStyle = 0
            .MaskColor = bCol
            Set MaskPicture = .Image
        End With
    Else
        UserControl.BackStyle = 1
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

Private Sub UserControl_InitProperties()
    'valeurs par défaut
    bNotOk2 = True
    With Me
        .BackColor = vbWhite '
        .BackStyle = Opaque
        .Caption = "Caption" '
        .Font = Ambient.Font '
        .ForeColor = vbBlack '
        .Enabled = True '
        .Value = False
        .Alignment = vbLeftJustify
        .UnRefreshControl = False
        .UseUnicode = False
    End With
    bNotOk2 = False
    Call UserControl_Paint  'refresh
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp, vbKeyLeft:
            Call SendKeys("+{Tab}")
        Case vbKeyDown, vbKeyRight:
            Call SendKeys("{Tab}")
        Case vbKeySpace
            Call ChangeValue
    End Select
    
    Call Refresh
    
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_LostFocus()
    bHasFocus = False: bNotOk = False
    Call UserControl_Paint
End Sub

Private Sub UserControl_Terminate()
    'vire le subclassing
    If OldProc Then Call SetWindowLong(UserControl.hWnd, GWL_WNDPROC, OldProc)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("BackColor", Me.BackColor, &HC0C0C0)
        Call .WriteProperty("BackStyle", Me.BackStyle, Opaque)
        Call .WriteProperty("Caption", Me.Caption, "Caption")
        Call .WriteProperty("Font", Me.Font, Ambient.Font)
        Call .WriteProperty("ForeColor", Me.ForeColor, vbBlack)
        Call .WriteProperty("Enabled", Me.Enabled, True)
        Call .WriteProperty("Value", Me.Value, False)
        Call .WriteProperty("Alignment", Me.Alignment, vbLeftJustify)
        Call .WriteProperty("UnRefreshControl", Me.UnRefreshControl, False)
        Call .WriteProperty("UseUnicode", Me.UseUnicode, False)
    End With
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    bNotOk2 = True
    With PropBag
        Me.BackColor = .ReadProperty("BackColor", &HC0C0C0)
        Me.BackStyle = .ReadProperty("BackStyle", Opaque)
        Me.Caption = .ReadProperty("Caption", "Caption")
        Set Me.Font = .ReadProperty("Font", Ambient.Font)
        Me.ForeColor = .ReadProperty("ForeColor", vbBlack)
        Me.Enabled = .ReadProperty("Enabled", True)
        Me.Value = .ReadProperty("Value", False)
        Me.Alignment = .ReadProperty("Alignment", vbLeftJustify)
        Me.UnRefreshControl = .ReadProperty("UnRefreshControl", False)
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
Attribute hDC.VB_Description = "Get the control hDc"
Public Property Get hWnd() As Long: hWnd = UserControl.hWnd: End Property
Attribute hWnd.VB_Description = "Handle of the control"
Public Property Get BackStyle() As BackStyleConstants: BackStyle = lBackStyle: End Property
Attribute BackStyle.VB_Description = "Use a transparent control or not"
Public Property Let BackStyle(BackStyle As BackStyleConstants): lBackStyle = BackStyle: UserControl.BackStyle = BackStyle: bNotOk = False: UserControl_Paint: End Property
Public Property Get Caption() As String: Caption = sCaption: End Property
Attribute Caption.VB_Description = "Caption to display"
Public Property Let Caption(Caption As String): sCaption = Caption: Call SetAccessKeys: bNotOk = False: UserControl_Paint: bNotOk = True: End Property
Public Property Get ForeColor() As OLE_COLOR: ForeColor = lForeColor: End Property
Attribute ForeColor.VB_Description = "Color of text"
Public Property Let ForeColor(ForeColor As OLE_COLOR): lForeColor = ForeColor: UserControl.ForeColor = ForeColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get BackColor() As OLE_COLOR: BackColor = bCol: End Property
Attribute BackColor.VB_Description = "Backcolor"
Public Property Let BackColor(BackColor As OLE_COLOR)
UserControl.BackColor = BackColor
bCol = BackColor: bNotOk = False: UserControl_Paint:
End Property
Public Property Get Font() As StdFont: Set Font = UserControl.Font: End Property
Attribute Font.VB_Description = "Text font"
Public Property Set Font(Font As StdFont): Set UserControl.Font = Font: bNotOk = False: UserControl_Paint: End Property
Public Property Get Enabled() As Boolean: Enabled = bEnable: End Property
Attribute Enabled.VB_Description = "Enable or not the control"
Public Property Let Enabled(Enabled As Boolean)
bEnable = Enabled: bNotOk = False: UserControl_Paint
End Property
Public Property Get Value() As CheckBoxConstants: Value = tVal: End Property
Attribute Value.VB_Description = "Value"
Attribute Value.VB_MemberFlags = "200"
Public Property Let Value(Value As CheckBoxConstants): tVal = Value: bNotOk = False: UserControl_Paint: End Property
Public Property Get Alignment() As AlignmentConstants: Alignment = tAlig: End Property
Attribute Alignment.VB_Description = "Text alignment"
Attribute Alignment.VB_MemberFlags = "40"
Public Property Let Alignment(Alignment As AlignmentConstants): tAlig = Alignment: bNotOk = False: UserControl_Paint: End Property
Public Property Get UnRefreshControl() As Boolean: UnRefreshControl = bUnRefreshControl: End Property
Attribute UnRefreshControl.VB_Description = "Prevent to refresh control"
Public Property Let UnRefreshControl(UnRefreshControl As Boolean): bUnRefreshControl = UnRefreshControl: End Property
Public Property Get UseUnicode() As Boolean: UseUnicode = bUnicode: End Property
Public Property Let UseUnicode(UseUnicode As Boolean): bUnicode = UseUnicode: bNotOk = False: UserControl_Paint: bNotOk = True: End Property


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
'MAJ du controle
'=======================================================
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh control"
Dim R As RECT
Dim yVal As Long
Dim st As Long

    If bUnRefreshControl Then Exit Sub
    
    '//on efface
    Call UserControl.Cls
    
    UserControl.MaskPicture = Nothing
    If bEnable Then
        UserControl.ForeColor = lForeColor
    Else
        UserControl.ForeColor = 10070188
    End If
    
    '//copnvertir les couleurs
    Call OleTranslateColor(bCol, 0, bCol)
    Call OleTranslateColor(lForeColor, 0, lForeColor)
    
    
    '//on va afficher l'image correspondant à l'état
    Call SplitIMGandShow
    
    '//on va tracer le rectangle de focus si on a le focus
    If bHasFocus Then Call UserControl_GotFocus
    
    '//affiche le texte
    yVal = (ScaleHeight - GetCharHeight * Screen.TwipsPerPixelY) / 2
    
    'définit une zone pour le texte
    Call SetRect(R, 20, yVal / Screen.TwipsPerPixelY, ScaleWidth / Screen.TwipsPerPixelX, ScaleHeight / Screen.TwipsPerPixelY)
    'dessine le texte
    If tAlig = vbLeftJustify Then
        st = DT_LEFT
    ElseIf tAlig = vbCenter Then
        st = DT_CENTER
    Else
        st = DT_RIGHT
    End If
    
    If bUnicode = False Then
        Call DrawText(UserControl.hDC, sCaption, Len(sCaption), R, st)
    Else
        Call DrawTextW(UserControl.hDC, StrPtr(sCaption), Len(sCaption), R, st)
    End If
    
    
    '//style
    If lBackStyle = [TRANSPARENT] Then
        'transparent
        With UserControl
            .BackStyle = 0
            .MaskColor = bCol
            Set MaskPicture = .Image
        End With
    Else
        UserControl.BackStyle = 1
    End If
    
    '//on refresh le control
    Call UserControl.Refresh
    
    bNotOk = True
End Sub

'=======================================================
'affiche une des 6 images en la découpant depuis l'image complète
'=======================================================
Private Sub SplitIMGandShow()
Dim SrcDC As Long
Dim SrcObj As Long
Dim y As Single
Dim lIMG As Long

    '0 rien
    '1 survol
    '2 enabled=false
    '3 value enable
    '4 value survol enable
    '5 enable=false OR gray
    
    If bEnable = False Then
        'grisé
        If tVal = vbUnchecked Then
            'pas checked
            lIMG = 2
        Else
            'checked et gris
            lIMG = 5
        End If
    Else
        'enabled=true
        If IsMouseIn Then
            'alors mouse survol
            If tVal = vbChecked Then
                'checked
                lIMG = 4
            ElseIf tVal = vbUnchecked Then
                'non checked
                lIMG = 1
            Else
                'gray
                lIMG = 5
            End If
        Else
            'pas de survol
            If tVal = vbChecked Then
                'checked
                lIMG = 3
            ElseIf tVal = vbUnchecked Then
                'non checked
                lIMG = 0
            Else
                'gray
                lIMG = 5
            End If
        End If
    End If
    
    'on découpe l'image correspondant à lIMG depuis Image1 et on blit
    'sur l'usercontrol
    SrcDC = CreateCompatibleDC(hDC)
    SrcObj = SelectObject(SrcDC, Image1.Picture)
    
    y = (ScaleHeight / Screen.TwipsPerPixelY - 13) / 2
    Call BitBlt(UserControl.hDC, 0, y, 13, 13, SrcDC, lIMG * 13, 0, SRCCOPY)

    Call DeleteDC(SrcDC)
    Call DeleteObject(SrcObj)
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
