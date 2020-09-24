Attribute VB_Name = "mdlTimerCallBack"
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

Public Timers As New Collection 'collection de tous les timers
'New pour instancier l'objet dès le début


'=======================================================
'function de callback pour le timer
'=======================================================
Public Sub TimerCallBackFunction(ByVal HWnd As Long, ByVal uMsg As Long, _
    ByVal IdEvent As Long, ByVal SysTime As Long)

Dim Tim As vkTimer  'contiendra LE timer qui appelle cette fonction de callback
Dim hTim As Long

    On Error Resume Next

    'récupère le pointeur sur l'objet ayant créé le timer
    hTim = CLng(Timers.Item("_" & CStr(IdEvent)))

    'on copie le Timer appelant (l'unique) dans la variable locale...
    Call CopyMemory(Tim, hTim, 4)  '4 octets

    'on appelle une sub du controle ayant créé le timer
    Call Tim.Raiser

    'on delete l'objet temporaire
    Call CopyMemory(Tim, CLng(0), 4)
End Sub


'Public Sub TimerCallBackFunction(ByVal uID As Long, ByVal uMsg As Long, _
'    ByVal dwUser As Long, ByVal dw1 As Long, ByVal dw2 As Long)
'
'Dim Tim As vkTimer  'contiendra LE timer qui appelle cette fonction de callback
'Dim hTim As Long
'
'    On Error Resume Next
'
'    'récupère le pointeur sur l'objet ayant créé le timer
'    hTim = CLng(Timers.Item("_" & CStr(uID)))
'
'    'on copie le Timer appelant (l'unique) dans la variable locale...
'    Call CopyMemory(Tim, hTim, 4)  '4 octets
'
'    'on appelle une sub du controle ayant créé le timer
'    Call Tim.Raiser
'
'    'on delete l'objet temporaire
'    Call CopyMemory(Tim, CLng(0), 4)
'
'End Sub

'=======================================================
'ajoute un timer à la liste
'=======================================================
Public Sub AddTimer(Obj As Long, ID As String)
    Call Timers.Add(Obj, ID)    'obj ==> pointeur sur vkTimer
End Sub

'=======================================================
'enlève un timer de la collection
'=======================================================
Public Sub RemoveTimer(ID As String)
Dim x As Long

    For x = 1 To Timers.Count
        If Timers.Item(x) = ID Then
            Timers.Remove (x)
            Exit For
        End If
    Next x
End Sub
