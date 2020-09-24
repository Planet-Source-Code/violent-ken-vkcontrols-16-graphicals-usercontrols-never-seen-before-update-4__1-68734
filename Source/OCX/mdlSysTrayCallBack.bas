Attribute VB_Name = "mdlSysTrayCallBack"
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

Public Trays As New Collection 'collection de tous les timers
'New pour instancier l'objet dès le début
Global OldOne As Long
Global msgTask As Long

'=======================================================
'ajoute un systray à la liste
'=======================================================
Public Sub AddVkTray(Obj As Long, ID As String)
    On Error GoTo YU
    Call Trays.Add(Obj, ID)    'obj ==> pointeur sur vkSysTray
YU:
End Sub

'=======================================================
'enlève un systray de la collection
'=======================================================
Public Sub RemoveVkTray(ID As String)
Dim x As Long

    For x = 1 To Trays.Count
        If Trays.Item(x) = ID Then
            Call Trays.Remove(x)
            Exit For
        End If
    Next x
End Sub

'=======================================================
'function de subclassing du container du controle Systray
'car un UserControl ne peux pas capter le message TaskBarCreated
'=======================================================
Public Function TrayContainerCallBackFunction(ByVal hwnd As Long, ByVal _
    uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
Dim Tim As vkSysTray    'contiendra LE controle qui appelle cette fonction de callback
Dim hTim As Long
Dim x As Long
    
    If uMsg = msgTask Then
        
        'pour tous les Items contenus dans la collection, on lance la procédure
        'RefreshAllIcons pour retracer les icones
        
        For x = 1 To Trays.Count
        
            'récupère le pointeur sur l'objet ayant créé le timer
            hTim = Trays.Item(x)
        
            If hTim <> 0 Then
    
                'on copie le vkSysTray dans la variable locale...
                Call CopyMemory(Tim, hTim, 4)  '4 octets
                
                'on appelle la procédure de rafraichissement des icones
                Call Tim.RefreshAllIcons
            
                'on delete l'objet temporaire
                Call CopyMemory(Tim, CLng(0), 4)
            End If
        Next x
        
    End If
    
    'alors on remet la WindowProc par défaut
    TrayContainerCallBackFunction = CallWindowProc(OldOne, hwnd, uMsg, wParam, lParam)

End Function

