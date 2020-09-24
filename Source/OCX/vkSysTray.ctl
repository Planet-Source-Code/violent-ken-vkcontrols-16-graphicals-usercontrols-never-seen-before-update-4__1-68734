VERSION 5.00
Begin VB.UserControl vkSysTray 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "vkSysTray.ctx":0000
   PropertyPages   =   "vkSysTray.ctx":0B0A
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "vkSysTray.ctx":0B1C
End
Attribute VB_Name = "vkSysTray"
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


Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'=======================================================
'VARIABLES PRIVEES
'=======================================================
Private mAsm(63) As Byte    'contient le code ASM
Private OldProc As Long     'adresse de l'ancienne window proc
Private objHwnd As Long     'handle de l'objet concerné

Private sBalloonTip As String
Private tPic As Picture
Private bIs() As Boolean
Private tIcon() As NOTIFYICONDATA
Private lTaskBarCreated As Long
Private bPreventFromCrash As Boolean


'=======================================================
'EVENTS
'=======================================================
Public Event MouseUp(Button As MouseButtonConstants, ID As Long)
Public Event MouseDown(Button As MouseButtonConstants, ID As Long)
Public Event MouseDblClick(Button As MouseButtonConstants, ID As Long)
Public Event MouseMove(Button As MouseButtonConstants, ID As Long)




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
    
    'appelle la gestion des events
    Call Raiser(uMsg, lParam, wParam, wParam)
        
    'appel de la routine standard pour les autres messages
    WindowProc = CallWindowProc(OldProc, hWnd, uMsg, wParam, lParam)
    
End Function

Private Sub UserControl_Initialize()
Dim Ofs As Long
Dim Ptr As Long
Dim Addr As Long
    
    'dimensionne le tableau des TrayIcons
    ReDim tIcon(0)
    ReDim bIs(0)
        
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
    Addr = Ofs
    MovB Ofs, &HA1                 'A1 20 20 40 00       mov         eax,RetVal
    MovL Ofs, VarPtr(mAsm(59))
    MovL Ofs, &H10C2               'C2 10 00             ret         10h
    
End Sub

Private Sub UserControl_InitProperties()
    With Me
        .BalloonTipString = vbNullString
        Set Me.Icon = Nothing
        .PreventFromCrash = True
    End With
End Sub

'=======================================================
'les events sont libérés dans cette procédure
'=======================================================
Private Sub Raiser(lParam As Long, uMsg As Long, wParam As Long, ID As Long)


    'Debug.Print uMsg & "  " & lParam & "  " & wParam & "  " & ID
    
    '/!\
    If uMsg = WM_LBUTTONDBLCLK Or uMsg = WM_LBUTTONDOWN Or uMsg = WM_LBUTTONUP Or uMsg _
        = WM_RBUTTONUP Or uMsg = WM_MBUTTONUP Or uMsg = WM_MBUTTONDOWN Or uMsg _
        = WM_RBUTTONDOWN Or uMsg = WM_MBUTTONDBLCLK Or uMsg = WM_RBUTTONDBLCLK Then
        
        'ceci permet de pouvoir cliquer à côté du menu en le faisant disparaitre
        'si l'user a prévu un popup menu
        Call SetForegroundWindow(GetParent(UserControl.hWnd))
    End If
    
    
    Select Case uMsg
        
        Case WM_LBUTTONDBLCLK
                RaiseEvent MouseDblClick(vbLeftButton, ID)
        Case WM_LBUTTONDOWN
                RaiseEvent MouseDown(vbLeftButton, ID)
        Case WM_LBUTTONUP
                RaiseEvent MouseUp(vbLeftButton, ID)
        Case WM_MBUTTONDBLCLK
                RaiseEvent MouseDblClick(vbMiddleButton, ID)
        Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, ID)
        Case WM_MBUTTONUP
                RaiseEvent MouseUp(vbMiddleButton, ID)
        Case WM_RBUTTONDBLCLK
                RaiseEvent MouseDblClick(vbRightButton, ID)
        Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, ID)
        Case WM_RBUTTONUP
                RaiseEvent MouseUp(vbRightButton, ID)
        Case WM_MOUSEMOVE
                RaiseEvent MouseMove(0, ID)
        Case lTaskBarCreated
                Call RefreshAllIcons
    End Select
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Me.BalloonTipString = .ReadProperty("BalloonTipString", vbNullString)
        Set Me.Icon = .ReadProperty("Icon", Nothing)
        Me.PreventFromCrash = .ReadProperty("PreventFromCrash", True)
    End With
    
    'lance hook et subclassing
    Call LaunchKeyMouseEvents
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("BalloonTipString", Me.BalloonTipString, vbNullString)
        Call .WriteProperty("Icon", Me.Icon, Nothing)
        Call .WriteProperty("PreventFromCrash", Me.PreventFromCrash, True)
    End With
End Sub

Private Sub UserControl_Resize()
    With UserControl
        .Width = 450
        .Height = 450
    End With
End Sub

Private Sub LaunchKeyMouseEvents()
                
    If Ambient.UserMode Then

        'change la WindowProc
        OldProc = SetWindowLong(UserControl.hWnd, GWL_WNDPROC, _
            VarPtr(mAsm(0)))    'pas de AddressOf aujourd'hui ;)
            
        'lance un hook sur le message TaskBarCreated (permet de récupérer le moment
        'ou la zone est recréée, pour retracer nos images, par exemple quand
        'explorer a crashé)
        If OldOne = 0 And bPreventFromCrash Then  'on ne fait QU'UNE fois le subclassing
            lTaskBarCreated = RegisterWindowMessage("TaskbarCreated")
            OldOne = SetWindowLong(UserControl.ContainerHwnd, GWL_WNDPROC, AddressOf TrayContainerCallBackFunction)
            msgTask = lTaskBarCreated
        End If
        
    End If
    
End Sub

'=======================================================
'désinstalle le sublclassing
'=======================================================
Private Sub UserControl_Terminate()
Dim x As Long

    'vire les subclassing
    If OldProc Then Call SetWindowLong(UserControl.hWnd, GWL_WNDPROC, OldProc)
    'If OldOne Then Call SetWindowLong(UserControl.ContainerHwnd, GWL_WNDPROC, OldOne)

    'on vire toutes les icones
    For x = 0 To UBound(tIcon())

        'enlève tous les Tray de la collection
        Call RemoveVkTray("_" & CStr(hWnd) & CStr(x))

        'enlève l'image
        Call Shell_NotifyIcon(NIM_DELETE, tIcon(x))

    Next x

    ReDim tIcon(0)
    Erase tIcon
End Sub


'=======================================================
'PROPERTIES
'=======================================================
Public Property Get IsInTray(ID As Long) As Boolean: On Error Resume Next: IsInTray = bIs(ID): End Property
Public Property Get BalloonTipString() As String: BalloonTipString = sBalloonTip: End Property
Public Property Let BalloonTipString(BalloonTipString As String): sBalloonTip = BalloonTipString: End Property
Public Property Get Icon() As Picture: Set Icon = tPic: End Property
Public Property Set Icon(NewPic As Picture): Set tPic = NewPic: End Property
Public Property Get PreventFromCrash() As Boolean: PreventFromCrash = bPreventFromCrash: End Property
Public Property Let PreventFromCrash(PreventFromCrash As Boolean): bPreventFromCrash = PreventFromCrash: End Property


'=======================================================
'PUBLIC FUNCTIONS & PROCEDURES
'=======================================================

'=======================================================
'ajoute dans le tray
'=======================================================
Public Sub AddToTray(Optional ByVal ID As Long = 0)

    'vérifie que l'ID est bon
    If ID < 0 Then Exit Sub
    
    If ID > UBound(tIcon()) Then
        'alors on ajoute une entée à ce tableau
        ReDim Preserve tIcon(ID)
        ReDim Preserve bIs(ID)
    End If
    
    bIs(ID) = True
    
    If tPic Is Nothing Then Exit Sub
    
    'on définit notre type perso pour le passer après en paramètre de l'API
    With tIcon(ID)
        .cbSize = Len(tIcon(ID))
        .hWnd = UserControl.hWnd
        .uID = ID
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = tPic.Handle
        .szTip = sBalloonTip & vbNullChar
        .uTimeout = 3000
        .dwInfoFlags = 1
    End With
    
    'on ajoute au TraySystem
    Call Shell_NotifyIcon(NIM_ADD, tIcon(ID))

    'on ajoute ce Tray à la collection
    Call AddVkTray(ObjPtr(Me), "_" & CStr(hWnd) & CStr(ID))   'pointeur sur CE vkSysTray
End Sub

'=======================================================
'enlève du tray
'=======================================================
Public Sub RemoveFromTray(Optional ByVal ID As Long = 0)
    
    'on vérifie l'ID
    If ID < 0 Or ID > UBound(tIcon()) Then Exit Sub
    
    'on vire du tray
    Call Shell_NotifyIcon(NIM_DELETE, tIcon(ID))
    
    'on vire de la collection
    Call RemoveVkTray("_" & CStr(hWnd) & CStr(ID))
    
    'on ne redimensionne QUE si on vient de virer le dernier Item<>0
    If UBound(tIcon()) = ID And ID > 0 Then
        ReDim Preserve tIcon(ID - 1)
        ReDim Preserve bIs(ID - 1)
    Else
        bIs(ID) = False
    End If

End Sub

'=======================================================
'refresh les icones
'=======================================================
Public Function RefreshAllIcons2()
Dim x As Long
    For x = 0 To UBound(tIcon())
        If bIs(x) Then
            'on ré-ajoute au TraySystem
            Call Shell_NotifyIcon(NIM_ADD, tIcon(x))
        End If
    Next x
End Function

'=======================================================
'c'est cette Function qui sera appellée pour recréer toutes les icones
'=======================================================
Friend Function RefreshAllIcons()
Dim Obj As Object
    
    On Error Resume Next
    
    'on refresh les icones de TOUS les vkSysTray du container
    For Each Obj In Parent.Controls
        If TypeOf Obj Is vkSysTray Then
            Call Obj.RefreshAllIcons2
        End If
    Next Obj
    
End Function
   


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
