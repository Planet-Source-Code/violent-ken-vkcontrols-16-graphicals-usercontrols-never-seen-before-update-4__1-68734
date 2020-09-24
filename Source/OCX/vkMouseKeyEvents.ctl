VERSION 5.00
Begin VB.UserControl vkMouseKeyEvents 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   MaskColor       =   &H00FFFFFF&
   MaskPicture     =   "vkMouseKeyEvents.ctx":0000
   Picture         =   "vkMouseKeyEvents.ctx":145E
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "vkMouseKeyEvents.ctx":28BC
End
Attribute VB_Name = "vkMouseKeyEvents"
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




'=======================================================
'EVENTS
'=======================================================
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Happens when subclassed control gets a keydown"
Public Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Happens when subclassed control gets a keypress"
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Happens when subclassed control gets a keyup"
Public Event MouseHover()
Attribute MouseHover.VB_Description = "Happens when mouse enters subclassed control"
Public Event MouseLeave()
Attribute MouseLeave.VB_Description = "Happens when mouse leaves subclassed control"
Public Event MouseWheel(Sens As Wheel_Sens)
Attribute MouseWheel.VB_Description = "Happens when subclassed control gets a wheel"
Public Event MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
Attribute MouseDown.VB_Description = "Happens when subclassed control gets a click"
Public Event MouseUp(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
Attribute MouseUp.VB_Description = "Happens when subclassed control gets a mouseup"
Public Event MouseDblClick(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
Attribute MouseDblClick.VB_Description = "Happens when subclassed control gets a dblclick"
Public Event MouseMove(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
Attribute MouseMove.VB_Description = "Happens when mouse moves on subclassed control"




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
    
        Case WM_KEYDOWN
                RaiseEvent KeyDown(CInt(wParam), GetShiftState)
        
        Case WM_KEYUP
                RaiseEvent KeyUp(CInt(wParam), GetShiftState)
        
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
        
    End Select
    
    'appel de la routine standard pour les autres messages
    WindowProc = CallWindowProc(OldProc, hWnd, uMsg, wParam, lParam)
    
End Function

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
    With UserControl
        .Width = 660
        .Height = 585
    End With
End Sub

'=======================================================
'lance le subclassing
'=======================================================
Public Sub LaunchKeyMouseEvents()
Attribute LaunchKeyMouseEvents.VB_Description = "Start subclassing of control"
    If Ambient.UserMode Then
        If objHwnd <> 0 Then
        
            
            OldProc = SetWindowLong(objHwnd, GWL_WNDPROC, _
                VarPtr(mAsm(0)))    'pas de AddressOf aujourd'hui ;)
                
            'prépare le terrain pour le mouse_over et mouse_leave
            With ET
                .cbSize = Len(ET)
                .hwndTrack = objHwnd
                .dwFlags = TME_LEAVE Or TME_HOVER
                .dwHoverTime = 1
            End With
            
            'démarre le tracking de l'entrée
            Call TrackMouseEvent(ET)
            
            'pas dedans par défaut
            IsMouseIn = False
        End If
    End If
End Sub

'=======================================================
'désinstalle le sublclassing
'=======================================================
Private Sub UserControl_Terminate()
        If OldProc And objHwnd Then _
            Call SetWindowLong(objHwnd, GWL_WNDPROC, OldProc)
End Sub




'=======================================================
'PROPERTIES
'=======================================================
Public Property Get ControlHwnd() As Long: ControlHwnd = objHwnd: End Property
Attribute ControlHwnd.VB_Description = "hWnd of the control you want to subclass"
Public Property Let ControlHwnd(ControlHwnd As Long)
    If OldProc <> 0 Then
        'alors déjà subclassing ==> on le termine
        Call UserControl_Terminate
    End If
    objHwnd = ControlHwnd
End Property



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
'renvoie l'état de l'argument Shift
'=======================================================
Private Function GetShiftState() As Integer
Dim intTemp As Integer

   If GetAsyncKeyState(vbKeyShift) Then intTemp = intTemp Or vbShiftMask
   If GetAsyncKeyState(vbKeyMenu) Then intTemp = intTemp Or vbAltMask
   If GetAsyncKeyState(vbKeyControl) Then intTemp = intTemp Or vbCtrlMask
   
   GetShiftState = intTemp
End Function
