VERSION 5.00
Begin VB.UserControl vkTimer 
   BackStyle       =   0  'Transparent
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3060
   InvisibleAtRuntime=   -1  'True
   MaskColor       =   &H00FFFFFF&
   MaskPicture     =   "vkTimer.ctx":0000
   Picture         =   "vkTimer.ctx":0F06
   ScaleHeight     =   2520
   ScaleWidth      =   3060
   ToolboxBitmap   =   "vkTimer.ctx":1E0C
End
Attribute VB_Name = "vkTimer"
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


Public Event Timer()
Attribute Timer.VB_Description = "Happens each [interval] ms"

'=======================================================
'VARIABLES PRIVEES
'=======================================================
Private bEnable As Boolean
Private lInterval As Long
Private hTimer As Long
Private bNotOk As Boolean


'=======================================================
'PROPERTIES
'=======================================================
Public Property Get Enabled() As Boolean: Enabled = bEnable: End Property
Attribute Enabled.VB_Description = "Enable or not the control"
Public Property Let Enabled(Enabled As Boolean)
If bEnable <> Enabled Then
    bEnable = Enabled: Refresh
End If
End Property
Public Property Get Interval() As Long: Interval = lInterval: End Property
Attribute Interval.VB_Description = "Interval of timer"
Public Property Let Interval(Interval As Long)
If Interval <> lInterval Then
    If Interval < 0 Then Interval = 0
    lInterval = Interval
    Call UserControl_Terminate
    Refresh
End If
End Property



'=======================================================
'USERCONTROL SUBS
'=======================================================
Private Sub UserControl_InitProperties()
    With Me
        .Enabled = False
        .Interval = 0
    End With
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Me.Enabled = .ReadProperty("Enabled", False)
        Me.Interval = .ReadProperty("Interval", 0)
    End With
End Sub
Private Sub UserControl_Resize()
    With UserControl
        .Width = 525
        .Height = 525
    End With
End Sub
Private Sub UserControl_Terminate()
'on delete timer et queue

    If hTimer Then
        
        'delete le timer
        'Call timeKillEvent(hTimer)
        Call KillTimer(0, hTimer)
    
        'on supprime le timer de la collection des timers
        Call RemoveTimer("_" & CStr(hTimer))
    End If
    
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Interval", Me.Interval, 0)
        Call .WriteProperty("Enabled", Me.Enabled, False)
    End With
End Sub

'=======================================================
'on lance ou non un timer
'=======================================================
Private Sub Refresh()

    On Error Resume Next

    If Not (Ambient.UserMode) Then Exit Sub
    
    If bEnable = False Or lInterval = 0 Then
        'on supprime le timer
        If hTimer Then
            Call KillTimer(0, hTimer)
            'Call timeKillEvent(hTimer)
            'l'enlève de la collection
            Call RemoveTimer("_" & CStr(hTimer))
        End If
    Else
        'création du timer
        'hTimer = timeSetEvent(lInterval, 0, AddressOf mdlTimerCallBack.TimerCallBackFunction, 0, _
            TIME_PERIODIC Or TIME_CALLBACK_FUNCTION)
        hTimer = SetTimer(0, 0, lInterval, AddressOf TimerCallBackFunction)
            
        'on ajoute un timer à la collection
        Call AddTimer(ObjPtr(Me), "_" & hTimer) 'pointeur sur CE vkTimer
    End If
End Sub

'=======================================================
'c'est ici que sera libéré l'event
'=======================================================
Friend Function Raiser()
    RaiseEvent Timer
End Function
