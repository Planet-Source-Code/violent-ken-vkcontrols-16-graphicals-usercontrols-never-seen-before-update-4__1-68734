VERSION 5.00
Begin VB.UserControl vkListBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3780
   PropertyPages   =   "vkListBox.ctx":0000
   ScaleHeight     =   3210
   ScaleWidth      =   3780
   ToolboxBitmap   =   "vkListBox.ctx":0042
   Begin vkUserContolsXP.vkVScrollPrivate VS 
      Height          =   2295
      Left            =   2400
      TabIndex        =   6
      Top             =   480
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   4048
      MouseInterval   =   50
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   5
      Left            =   600
      Picture         =   "vkListBox.ctx":0354
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   4
      Left            =   240
      Picture         =   "vkListBox.ctx":059E
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   3
      Left            =   1320
      Picture         =   "vkListBox.ctx":07E8
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   2
      Left            =   960
      Picture         =   "vkListBox.ctx":0A32
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   1
      Left            =   600
      Picture         =   "vkListBox.ctx":0C7C
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   0
      Left            =   240
      Picture         =   "vkListBox.ctx":0EC6
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "vkListBox"
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

Private bDisplayBorder As Boolean
Private lBackColor As OLE_COLOR
Private bEnable As Boolean
Private lListCount As Long
'Private lHeight() As Long
Private bSelected() As Boolean
Private bChecked() As Boolean
Private picCol As Collection
Private cFile As clsFile
Private bMultiSelect As Boolean
Private lNewIndex As Long
Private lSelCount As Long
Private bStyleCheckBox As Boolean
Private lTopIndex As Long
Private bNotOk As Boolean
Private bNotOk2 As Boolean
Private bUnRefreshControl As Boolean
Private FilePics As clsFastCollection
Private lForeColor As OLE_COLOR
Private lBorderColor As OLE_COLOR
Private bSorted As SortOrder
Private lCheckCount As Long
Private tAlig As AlignmentConstants
Private lSelColor As Long
Private lPrevSel As Long
Private vsPushed As Boolean
Private MouseItemIndex As Long
Private lFullRowSelect As Boolean
Private lBorderSelColor  As OLE_COLOR
Private tmpMouseItemIndex As Long
Private Col As clsFastCollection
Private zNumber As Long
Private bVSvisible As Boolean
Private tScroll As New vkPrivateScroll
Private bAcceptAutoSort As Boolean
Private tListType As ListBoxType
Private sPath As String
Private bOkToAddFile As Boolean
Private bDisplayFileIcons As Boolean
Private bDisplayEntirePath As Boolean
Private bUseDefautItemSettings As Boolean
Private lIconSize As IconSize
Private bShowHiddenFiles As Boolean
Private sPattern As String
Private bShowSystemFiles As Boolean
Private bShowReadOnlyFiles As Boolean
Private bUnicode As Boolean


'=======================================================
'EVENTS
'=======================================================
Public Event ItemClick(Item As vkListItem)
Public Event ItemDblClick(Item As vkListItem)
Public Event ItemUnCheck(Item As vkListItem)
Public Event ItemChek(Item As vkListItem)
Public Event Scroll()
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
Attribute WindowProc.VB_Description = "Internal proc for subclassing"
Attribute WindowProc.VB_MemberFlags = "40"
Dim iControl As Integer
Dim iShift As Integer
Dim z As Long
Dim x As Long
Dim y As Long
Dim s As Long
Dim e As Long
Dim F As Long, z3 As Long


    Select Case uMsg
        
        Case WM_LBUTTONDBLCLK
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
                'détermine quel Item est sélectionné
                s = 0   'hauteur temporaire
                On Error Resume Next
                For z = lTopIndex To lListCount
                    s = s + Col.Item(z).Height
                    If s > y Then e = z: Exit For
                Next z
                
                MouseItemIndex = e
                
                'permet de cocher les checks si jamais on double-click
                If bStyleCheckBox Then
                    bChecked(MouseItemIndex) = Not (bChecked(MouseItemIndex))
                    Call Refresh    'change l'état des images
                    If bChecked(MouseItemIndex) Then
                        RaiseEvent ItemChek(Col.Item(MouseItemIndex))
                    Else
                        RaiseEvent ItemUnCheck(Col.Item(MouseItemIndex))
                    End If
                 End If
    
                RaiseEvent ItemDblClick(Col.Item(e))
                RaiseEvent MouseDblClick(vbLeftButton, iShift, iControl, x, y)
        Case WM_LBUTTONDOWN
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
                'détermine quel Item est sélectionné
                s = 0   'hauteur temporaire
                On Error Resume Next
                For z = lTopIndex To lListCount
                    s = s + Col.Item(z).Height
                    If s > y Then e = z: Exit For
                Next z
                
                MouseItemIndex = e
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
                
                'détermine quel Item est sélectionné
                s = 0   'hauteur temporaire
                On Error Resume Next
                For z = lTopIndex To lListCount
                    s = s + Col.Item(z).Height
                    If s > y Then e = z: Exit For
                Next z
                
                MouseItemIndex = e
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

                If bStyleCheckBox Then
                    On Error Resume Next
                    For F = lTopIndex To lListCount
                        e = e + Col.Item(F).Height
                        If e >= Height - 50 Then
                            z3 = z
                            Exit For
                        End If
                        z = z + 1
                    Next F
            
                    Call SplitIMGandShow(z)
                End If
            
            End If
        Case WM_MOUSELEAVE
        
            If bStyleCheckBox Then
                On Error Resume Next
                For F = lTopIndex To lListCount
                    e = e + Col.Item(F).Height
                    If e >= Height - 50 Then
                        z3 = z
                        Exit For
                    End If
                    z = z + 1
                Next F
        
                MouseItemIndex = -1: Call SplitIMGandShow(z)
            End If
            
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
                
                'détermine quel Item est sélectionné
                s = 0   'hauteur temporaire
                On Error Resume Next
                For z = lTopIndex To lListCount
                    s = s + Col.Item(z).Height
                    If s > y Then e = z: Exit For
                Next z
                
                MouseItemIndex = e
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
    
    bNotOk = True
    bNotOk2 = True
    ReDim bSelected(1)
    ReDim bChecked(1)
    ReDim tPic(1)
    lListCount = 1
    
    'instancie la collection
    Set Col = New clsFastCollection
    Set picCol = New Collection
    
    'instancie la classe File et la collection d'images
    Set cFile = New clsFile
    Set FilePics = New clsFastCollection
    
    'créé un controle dynamique
    'Set VS = Controls.Add("vkUserContolsXP.vkVScroll", "VS")
    'VS.MyExtender.Visible = True
    
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
    
    'initialise le VS
    bUnRefreshControl = True
    With VS
        .UnRefreshControl = True
        .Max = 1
        .Min = 1
        .Value = 1
        .Visible = True
        .UnRefreshControl = False
        'Call .Refresh
    End With
    bUnRefreshControl = False
End Sub

Private Sub UserControl_InitProperties()
    'valeurs par défaut
    bNotOk2 = True
    With Me
        .BackColor = &HFFFFFF
        .BorderColor = 12937777
        .DisplayBorder = True
        .Pattern = "*.*"
        .Enabled = True
        .ForeColor = vbBlack
        .MultiSelect = True
        .Sorted = False
        .StyleCheckBox = False
        .UnRefreshControl = False
        .ListIndex = -1
        .DisplayVScroll = True
        .Alignment = vbLeftJustify
        .SelColor = 16768444
        .Font = Ambient.Font
        .FullRowSelect = True
        .BorderSelColor = 16419097
        .TopIndex = 1
        .AcceptAutoSort = False
        .ListType = SimpleList
        .Path = App.Path
        .DisplayFileIcons = True
        .UseDefautItemSettings = False
        .DisplayEntirePath = False
        .IconSize = Size16
        .ShowHiddenFiles = True
        .ShowReadOnlyFiles = True
        .ShowSystemFiles = True
        .UseUnicode = False
    End With
    bNotOk2 = False
    bNotOk = True: Call UserControl_Paint 'refresh
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
Private Sub vkVScroll1_Scroll()
    RaiseEvent Scroll
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'sélection d'un item
Dim z As Long
Dim s As Long
Dim e As Long

    On Error Resume Next
    
    If bEnable = False Then Exit Sub
    
    'si dans la zone gauche et que style=checkboxes ==> on checke
    If bStyleCheckBox Then
        If x < 255 Then
            bChecked(MouseItemIndex) = Not (bChecked(MouseItemIndex))
            If bChecked(MouseItemIndex) Then
                RaiseEvent ItemChek(Col.Item(MouseItemIndex))
            Else
                RaiseEvent ItemUnCheck(Col.Item(MouseItemIndex))
            End If
        End If
    End If

    'détermine quel Item est sélectionné
    s = 0   'hauteur temporaire
    For z = lTopIndex To lListCount
        s = s + Col.Item(z).Height
        If s > y Then e = z: Exit For
    Next z
    
    RaiseEvent ItemClick(Col.Item(e))
    
    If bMultiSelect = False Then
        'déselectionne tout
        Call UnSelectAll(False)
    Else
        'alors on teste en fonction du Shift
        If (Shift And vbShiftMask) = vbShiftMask Then
            'on sélectionne tout entre lPrevSel et e-1
            Dim o As Boolean
            If e - 1 > lPrevSel Then
                o = bSelected(lPrevSel)
            End If
            For s = e To lPrevSel Step IIf(e - 1 < lPrevSel, 1, -1)
                'Col.Item(s).Selected = True
                bSelected(s) = True
            Next s
            If e - 1 > lPrevSel Then
                'on supprime le premier(terme correctif)
                'Col.Item(lPrevSel).Selected = o
                bSelected(lPrevSel) = o
            End If
        ElseIf (Shift And vbCtrlMask) = vbCtrlMask Then
            'on permute le sélectionné et on touche pas au reste
            'Col.Item(e).Selected = Not (Col.Item(e).Selected)
            bSelected(e) = Not (bSelected(e))
            Call Refresh
            Exit Sub    'évite de revenir à selected(e)=true
        Else
            'déselectionne tout
            Call UnSelectAll(False)
        End If
    End If
        
    
    'alors si un élément est sélectionné
    If e Then
        'Col.Item(e).Selected = True
        bSelected(e) = True
        Call Refresh
    End If
    
    'sauvegarde le dernier Item sauvegardé
    lPrevSel = e - 1
        
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim z As Long
Dim z2 As Long
Dim z3 As Long
Dim e As Long
Dim m As Long
Dim F As Long
    
    If bStyleCheckBox = False Then Exit Sub
    
    On Error Resume Next

    'on détermine quel Item est survolé
    z2 = -1
    For F = lTopIndex To lListCount
        e = e + Col.Item(F).Height
        If e > y Then
            If z2 = -1 Then z2 = z: m = e - Col.Item(F).Height
        End If
        If e >= Height - 50 Then
            z3 = z
            Exit For
        End If
        z = z + 1
    Next F
    
    'si pas suffisemment d'items pour remplir la vue
    'alors le nombre d'affichés = listcount
    If z3 = 0 Then z3 = ListCount
    
    'récupère l'Item survolé
    MouseItemIndex = lTopIndex + z2
    
    'redessine les images si nécessaire (item survolé différent)
    If MouseItemIndex <> tmpMouseItemIndex Then Call SplitIMGandShow(z3)
    
    'sauvegarde les bornes (en height) de l'item survolé
    tmpMouseItemIndex = MouseItemIndex
    
End Sub

Private Sub UserControl_Terminate()
    'vire le subclassing
    If OldProc Then Call SetWindowLong(UserControl.hWnd, GWL_WNDPROC, OldProc)
    'on clear la collection
    Call Col.Clear
    Set Col = Nothing
    Set tScroll = Nothing
    Set cFile = Nothing
    Set FilePics = Nothing
    Set picCol = Nothing
    'on efface les tableaux
    Erase bSelected
    Erase bChecked
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    bNotOk2 = True
    With PropBag
        Call .WriteProperty("Font", Me.Font, Ambient.Font)
        Call .WriteProperty("BackColor", Me.BackColor, &HFFFFFF)
        Call .WriteProperty("BorderColor", Me.BorderColor, 12937777)
        Call .WriteProperty("DisplayBorder", Me.DisplayBorder, True)
        Call .WriteProperty("Enabled", Me.Enabled, True)
        Call .WriteProperty("ForeColor", Me.ForeColor, vbBlack)
        Call .WriteProperty("MultiSelect", Me.MultiSelect, True)
        Call .WriteProperty("Sorted", Me.Sorted, True)
        Call .WriteProperty("StyleCheckBox", Me.StyleCheckBox, False)
        Call .WriteProperty("UnRefreshControl", Me.UnRefreshControl, False)
        Call .WriteProperty("ListIndex", Me.ListIndex, -1)
        Call .WriteProperty("DisplayVScroll", Me.DisplayVScroll, True)
        Call .WriteProperty("Alignment", Me.Alignment, vbLeftJustify)
        Call .WriteProperty("SelColor", Me.SelColor, 16768444)
        Call .WriteProperty("FullRowSelect", Me.FullRowSelect, True)
        Call .WriteProperty("BorderSelColor", Me.BorderSelColor, 16419097)
        Call .WriteProperty("TopIndex", Me.TopIndex, 1)
        'Call .WriteProperty("VScroll", Me.VScroll, tScroll)
        Call .WriteProperty("UseDefautItemSettings", Me.UseDefautItemSettings, False)
        Call .WriteProperty("AcceptAutoSort", Me.AcceptAutoSort, False)
        Call .WriteProperty("ListType", Me.ListType, SimpleList)
        Call .WriteProperty("Path", Me.Path, App.Path)
        Call .WriteProperty("DisplayFileIcons", Me.DisplayFileIcons, True)
        Call .WriteProperty("DisplayEntirePath", Me.DisplayEntirePath, False)
        Call .WriteProperty("IconSize", Me.IconSize, Size16)
        Call .WriteProperty("ShowHiddenFiles", Me.ShowHiddenFiles, True)
        Call .WriteProperty("ShowReadOnlyFiles", Me.ShowReadOnlyFiles, True)
        Call .WriteProperty("ShowSystemFiles", Me.ShowSystemFiles, True)
        Call .WriteProperty("Pattern", Me.Pattern, "*.*")
        Call .WriteProperty("UseUnicode", Me.UseUnicode, False)
    End With
    bNotOk2 = False
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    bNotOk2 = True
    With PropBag
        Me.BackColor = .ReadProperty("BackColor", &HFFFFFF)
        Me.BorderColor = .ReadProperty("BorderColor", 12937777)
        Me.DisplayBorder = .ReadProperty("DisplayBorder", True)
        Me.Enabled = .ReadProperty("Enabled", True)
        Me.ForeColor = .ReadProperty("ForeColor", vbBlack)
        Me.MultiSelect = .ReadProperty("MultiSelect", True)
        Me.Sorted = .ReadProperty("Sorted", True)
        Me.StyleCheckBox = .ReadProperty("StyleCheckBox", False)
        Set Me.Font = .ReadProperty("Font", Ambient.Font)
        Me.UnRefreshControl = .ReadProperty("UnRefreshControl", False)
        Me.ListIndex = .ReadProperty("ListIndex", -1)
        Me.DisplayVScroll = .ReadProperty("DisplayVScroll", True)
        Me.Alignment = .ReadProperty("Alignment", vbLeftJustify)
        Me.SelColor = .ReadProperty("SelColor", 16768444)
        Me.FullRowSelect = .ReadProperty("FullRowSelect", True)
        Me.BorderSelColor = .ReadProperty("BorderSelColor", 16419097)
        Me.TopIndex = .ReadProperty("TopIndex", 1)
        'Me.VScroll.UnRefreshControl = True
        Me.VScroll = .ReadProperty("VScroll", tScroll)
        Me.AcceptAutoSort = .ReadProperty("AcceptAutoSort", False)
        Me.ListType = .ReadProperty("ListType", SimpleList)
        Me.Path = .ReadProperty("Path", App.Path)
        Me.DisplayFileIcons = .ReadProperty("DisplayFileIcons", True)
        Me.DisplayEntirePath = .ReadProperty("DisplayEntirePath", False)
        Me.IconSize = .ReadProperty("IconSize", Size16)
        Me.UseDefautItemSettings = .ReadProperty("UseDefautItemSettings", False)
        Me.ShowHiddenFiles = .ReadProperty("ShowHiddenFiles", True)
        Me.ShowReadOnlyFiles = .ReadProperty("ShowReadOnlyFiles", True)
        Me.ShowSystemFiles = .ReadProperty("ShowSystemFiles", True)
        Me.Pattern = .ReadProperty("Pattern", "*.*")
        Me.UseUnicode = .ReadProperty("UseUnicode", False)
    End With
    bNotOk2 = False
    If tListType = SimpleList Then
        Call Refresh  'refresh
    Else
        Call AddFileToList
    End If
    
    'le bon endroit pour lancer le subclassing
    Call LaunchKeyMouseEvents
    If Ambient.UserMode Then
        Call VS.LaunchKeyMouseEvents    'subclasse également le VS
    End If
End Sub
Private Sub UserControl_Resize()
    If Height < 800 Then Height = 800
    With VS
        .UnRefreshControl = True
        .Width = 255
        .Top = 0
        .Left = Width - 255
        .Height = Height
        .UnRefreshControl = False
        'Call .Refresh
    End With
    bNotOk = False: Call UserControl_Paint 'refresh
    
    'bon ben lors du premier resize, on considère que le controle est près
    'pour pouvoir être remplit par les items File/Folder
    If bOkToAddFile = False Then
        bOkToAddFile = True
        Call AddFileToList
    End If
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

'//Propriétés sur le Scroll
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
        '.Width = VScroll.Width
        '.Left = Width - .Width
        .UnRefreshControl = False
    End With
    Call VS.Refresh
    bNotOk = False: Call UserControl_Paint
End Property

'//Proptiétés normales
Public Property Get hDC() As Long: hDC = UserControl.hDC: End Property
Public Property Get hWnd() As Long: hWnd = UserControl.hWnd: End Property
Public Property Get SelColor() As OLE_COLOR: SelColor = lSelColor: End Property
Public Property Let SelColor(SelColor As OLE_COLOR): lSelColor = SelColor: UserControl.ForeColor = ForeColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get ForeColor() As OLE_COLOR: ForeColor = lForeColor: End Property
Public Property Let ForeColor(ForeColor As OLE_COLOR): lForeColor = ForeColor: UserControl.ForeColor = ForeColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get BackColor() As OLE_COLOR: BackColor = lBackColor: End Property
Public Property Let BackColor(BackColor As OLE_COLOR): UserControl.BackColor = BackColor: lBackColor = BackColor: End Property
Public Property Get Font() As StdFont: Set Font = UserControl.Font: End Property
Public Property Set Font(Font As StdFont): Set UserControl.Font = Font: bNotOk = False: UserControl_Paint: End Property
Public Property Get Enabled() As Boolean: Enabled = bEnable: End Property
Public Property Let Enabled(Enabled As Boolean): VS.Enabled = Enabled: bEnable = Enabled: bNotOk = False: UserControl_Paint: End Property
Public Property Get DisplayBorder() As Boolean: DisplayBorder = bDisplayBorder: End Property
Public Property Let DisplayBorder(DisplayBorder As Boolean): bDisplayBorder = DisplayBorder: bNotOk = False: UserControl_Paint: End Property
Public Property Get BorderColor() As OLE_COLOR: BorderColor = lBorderColor: End Property
Public Property Let BorderColor(BorderColor As OLE_COLOR): lBorderColor = BorderColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get List(Index As Long) As String: On Error Resume Next: List = Col.Item(Index).Text: End Property
Public Property Let List(Index As Long, List As String): On Error Resume Next: Col.Item(Index).Text = List: bNotOk = False: UserControl_Paint: End Property
Public Property Get ListCount() As Long: ListCount = lListCount - 1: End Property
Public Property Get ListIndex() As Long: ListIndex = MouseItemIndex: End Property
Public Property Let ListIndex(ListIndex As Long): MouseItemIndex = ListIndex: bNotOk = False: UserControl_Paint: End Property
Public Property Get MultiSelect() As Boolean: MultiSelect = bMultiSelect: End Property
Public Property Let MultiSelect(MultiSelect As Boolean): bMultiSelect = MultiSelect: End Property
Public Property Get NewIndex() As Long: NewIndex = lNewIndex: End Property
Public Property Get SelCount() As Long: SelCount = lSelCount: End Property
Public Property Get Selected(Index As Long) As Boolean: On Error Resume Next: Selected = bSelected(Index): End Property
Public Property Get Checked(Index As Long) As Boolean: On Error Resume Next: Checked = bChecked(Index): End Property
Public Property Let Selected(Index As Long, Selected As Boolean): On Error Resume Next: bSelected(Index) = Selected: End Property
Public Property Let Checked(Index As Long, Checked As Boolean): On Error Resume Next: bChecked(Index) = Checked: End Property
Public Property Get Sorted() As SortOrder: Sorted = bSorted: End Property
Public Property Let Sorted(Sorted As SortOrder)
bSorted = Sorted: bNotOk = False: Call Sort(bSorted): UserControl_Paint
End Property
Public Property Get TopIndex() As Long: TopIndex = lTopIndex: End Property
Public Property Let TopIndex(TopIndex As Long): lTopIndex = TopIndex: bNotOk = False: UserControl_Paint: End Property
Public Property Get StyleCheckBox() As Boolean: StyleCheckBox = bStyleCheckBox: End Property
Public Property Let StyleCheckBox(StyleCheckBox As Boolean): bStyleCheckBox = StyleCheckBox: bNotOk = False: UserControl_Paint: End Property
Public Property Get Item(Index As Long) As vkListItem
On Error Resume Next: Set Item = Col.Item(Index)
Item.Checked = bChecked(Index)
Item.Selected = bSelected(Index)
End Property
Public Property Let Item(Index As Long, Item As vkListItem)
On Error Resume Next
Set Col.Item(Index) = Item
bSelected(Index) = Item.Selected
bChecked(Index) = Item.Checked
'lHeight(Index) = Item.Height
bNotOk = False: UserControl_Paint
End Property
Public Property Get UnRefreshControl() As Boolean: UnRefreshControl = bUnRefreshControl: End Property
Attribute UnRefreshControl.VB_Description = "Prevent to refresh control"
Public Property Let UnRefreshControl(UnRefreshControl As Boolean): bUnRefreshControl = UnRefreshControl: End Property
Public Property Get DisplayVScroll() As Boolean: DisplayVScroll = bVSvisible: End Property
Public Property Let DisplayVScroll(DisplayVScroll As Boolean)
bVSvisible = DisplayVScroll
VS.Visible = bVSvisible
bNotOk = False: UserControl_Paint
End Property
Public Property Get CheckCount() As Long: CheckCount = lCheckCount: End Property
Public Property Get Alignment() As AlignmentConstants: Alignment = tAlig: End Property
Public Property Let Alignment(Alignment As AlignmentConstants): tAlig = Alignment: bNotOk = False: UserControl_Paint: End Property
Public Property Get ListItems() As clsFastCollection: On Error Resume Next: Set ListItems = Col: End Property
Public Property Let ListItems(ListItems As clsFastCollection): On Error Resume Next: Set Col = ListItems: bNotOk = False: UserControl_Paint: End Property
Public Property Get FullRowSelect() As Boolean: FullRowSelect = lFullRowSelect: End Property
Public Property Let FullRowSelect(FullRowSelect As Boolean): lFullRowSelect = FullRowSelect: bNotOk = False: UserControl_Paint: End Property
Public Property Get BorderSelColor() As OLE_COLOR: BorderSelColor = lBorderSelColor: End Property
Public Property Let BorderSelColor(BorderSelColor As OLE_COLOR): lBorderSelColor = BorderSelColor: UserControl.ForeColor = ForeColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get AcceptAutoSort() As Boolean: AcceptAutoSort = bAcceptAutoSort: End Property
Attribute AcceptAutoSort.VB_MemberFlags = "40"
Public Property Let AcceptAutoSort(AcceptAutoSort As Boolean): bAcceptAutoSort = AcceptAutoSort
If bAcceptAutoSort Then Call Sort(bSorted)
End Property
Public Property Get ListType() As ListBoxType: ListType = tListType: End Property
Public Property Let ListType(ListType As ListBoxType): tListType = ListType
bNotOk = False: Call AddFileToList
End Property
Public Property Get Path() As String: Path = sPath: End Property
Public Property Let Path(Path As String): sPath = Path
bNotOk = False: Call AddFileToList
End Property
Public Property Get DisplayFileIcons() As Boolean: DisplayFileIcons = bDisplayFileIcons: End Property
Public Property Let DisplayFileIcons(DisplayFileIcons As Boolean): bDisplayFileIcons = DisplayFileIcons
bNotOk = False: Call AddFileToList
End Property
Public Property Get DisplayEntirePath() As Boolean: DisplayEntirePath = bDisplayEntirePath: End Property
Public Property Let DisplayEntirePath(DisplayEntirePath As Boolean): bDisplayEntirePath = DisplayEntirePath
bNotOk = False: Call AddFileToList
End Property
Public Property Get UseDefautItemSettings() As Boolean: UseDefautItemSettings = bUseDefautItemSettings: End Property
Public Property Let UseDefautItemSettings(UseDefautItemSettings As Boolean): bUseDefautItemSettings = UseDefautItemSettings: bNotOk = False: UserControl_Paint: End Property
Public Property Get IconSize() As IconSize: IconSize = lIconSize: End Property
Public Property Let IconSize(IconSize As IconSize): lIconSize = IconSize: Call AddFileToList: End Property
Public Property Get ShowHiddenFiles() As Boolean: ShowHiddenFiles = bShowHiddenFiles: End Property
Public Property Let ShowHiddenFiles(ShowHiddenFiles As Boolean): bShowHiddenFiles = ShowHiddenFiles: bNotOk = False: AddFileToList: End Property
Public Property Get ShowSystemFiles() As Boolean: ShowSystemFiles = bShowSystemFiles: End Property
Public Property Let ShowSystemFiles(ShowSystemFiles As Boolean): bShowSystemFiles = ShowSystemFiles: bNotOk = False: AddFileToList: End Property
Public Property Get ShowReadOnlyFiles() As Boolean: ShowReadOnlyFiles = bShowReadOnlyFiles: End Property
Public Property Let ShowReadOnlyFiles(ShowReadOnlyFiles As Boolean): bShowReadOnlyFiles = ShowReadOnlyFiles: bNotOk = False: AddFileToList: End Property
Public Property Get Pattern() As String: Pattern = sPattern: End Property
Public Property Let Pattern(Pattern As String)
sPattern = Pattern
bNotOk = False: Call AddFileToList
End Property
Public Property Get UseUnicode() As Boolean: UseUnicode = bUnicode: End Property
Public Property Let UseUnicode(UseUnicode As Boolean): bUnicode = UseUnicode: bNotOk = False: UserControl_Paint: bNotOk = True: End Property


Private Sub UserControl_Paint()

    If bNotOk Or bNotOk2 Or bUnRefreshControl Then Exit Sub     'pas prêt à peindre
    
    If tListType = SimpleList Then
        Call Refresh
    Else
        Call AddFileToList
    End If
End Sub








'=======================================================
'PUBLIC SUBS
'=======================================================
'=======================================================
'ajoute un objet à la liste des objets
'=======================================================
Public Sub AddItem(Optional ByVal Caption As String, Optional ByVal Item As _
    vkListItem, Optional ByVal Key As String, Optional ByVal Index As Long = -1)
    
Dim tIt As vkListItem
Dim bOk As Boolean
    
    lListCount = lListCount + 1
        
    'redimensionne les tableaux avec le nombre d'items de la liste
    ReDim Preserve bChecked(lListCount - 1)
    ReDim Preserve bSelected(lListCount - 1)
    'ReDim Preserve lHeight(lListCount - 1)
    
    If Item Is Nothing Then
        'alors on créé un nouvel Item dont on définit les prop par défaut
        Set tIt = New vkListItem
        With tIt
            .BackColor = lBackColor
            .Checked = False
            .Font = UserControl.Font
            .ForeColor = lForeColor
            .Key = Key
            .Selected = False
            .Text = Caption
            .Height = TextHeight(.Text) + 50
            .Alignment = tAlig
            .SelColor = lSelColor
            .BorderSelColor = lBorderSelColor
        End With
        
        If Index = -1 Then
            tIt.Index = Col.Count + 1
            'lHeight(lListCount - 1) = tIt.Height
            Call Col.Add(tIt)
        Else
            tIt.Index = Index
            'lHeight(Index) = tIt.Height
            Call Col.Add(tIt, Index)
        End If
        
        lNewIndex = tIt.Index
        
    Else
        'on ajoute l'item passé en paramètre
        If Index = -1 Then
            Item.Index = lListCount - 1
            bSelected(Item.Index) = Item.Selected
            bChecked(Item.Index) = Item.Checked
            'lHeight(lListCount - 1) = Item.Height
            Call Col.Add(Item)
        Else
            bSelected(Index) = Item.Selected
            bChecked(Index) = Item.Checked
            'lHeight(Index) = Item.Height
            Call Col.Add(Item, Index)
        End If
        
        lNewIndex = Item.Index
        
    End If
    
    
    With VS
        .UnRefreshControl = True
        .Max = lListCount
        .UnRefreshControl = False
    End With
    
    'on trie à nouveau
    If bAcceptAutoSort Then bNotOk = False: Call Sort(bSorted)
    
    'on refresh
    Call Refresh
End Sub

'=======================================================
'efface tous les objets de la liste
'=======================================================
Public Sub Clear()
Dim x As Long
    
    'efface les tableau
    ReDim bSelected(1)
    ReDim bChecked(1)
    'ReDim lHeight(1)
    
    'on vide la collection...
    Call Col.Clear
    
    lListCount = 1
    lSelCount = 0
    lCheckCount = 0

    VS.Max = 1
    
    'refresh
    Call Refresh
End Sub

'=======================================================
'inverse la sélection
'=======================================================
Public Sub InvertSelection()
Dim x As Long
Dim y As Long

    'inverse le contenu du tableau
    For x = 1 To lListCount - 1
        bSelected(x) = Not (bSelected(x))
        If bSelected(x) Then y = y + 1
    Next x
    
    lSelCount = y
    
    'refresh
    Call Refresh
    
End Sub

'=======================================================
'inverse les cases cochées
'=======================================================
Public Sub InvertChecks()
Dim x As Long
Dim y As Long

    'inverse le contenu du tableau
    For x = 1 To lListCount - 1
        bChecked(x) = Not (bChecked(x))
        If bChecked(x) Then y = y + 1
    Next x
    
    lCheckCount = y
    
    'refresh
    Call Refresh
    
End Sub

'=======================================================
'enlève un item de la liste
'=======================================================
Public Sub RemoveItem(ByVal Index As Long)
    
    'vire l'item
    If bChecked(Index) Then
        lCheckCount = lCheckCount - 1
    End If
    If bSelected(Index) Then
        lSelCount = lSelCount - 1
    End If
    
    Call Col.Remove(Index)
    
    'on redimensionne les tableaux
    lListCount = lListCount - 1
    If lListCount < 1 Then lListCount = 1
    ReDim Preserve bChecked(lListCount)
    ReDim Preserve bSelected(lListCount)
    'ReDim Preserve lHeight(lListCount)
    
    VS.Max = lListCount - 1
    
    'refresh
    Call Refresh
    
End Sub

'=======================================================
'sélectionne tout
'=======================================================
Public Sub SelectAll()
Dim x As Long

    'remplit le contenu du tableau
    For x = 1 To lListCount - 1
        bSelected(x) = True
    Next x
    
    lSelCount = lListCount
    
    'refresh
    Call Refresh
    
End Sub

'=======================================================
'ne sélectionne rien
'=======================================================
Public Sub UnSelectAll(Optional ByVal RefreshControl As Boolean = True)
Dim x As Long
    
    'remplit le contenu du tableau
    ReDim bSelected(lListCount)
    
    lSelCount = 0
    
    'refresh
    If RefreshControl Then Call Refresh
    
End Sub

'=======================================================
'checke tout
'=======================================================
Public Sub CheckAll()
Dim x As Long

    'remplit le contenu du tableau
    For x = 1 To lListCount - 1
        bChecked(x) = True
    Next x
    
    lCheckCount = lListCount
    
    'refresh
    Call Refresh
    
End Sub

'=======================================================
'trie la collection
'=======================================================
Public Sub SortItems(ByVal SortType As SortOrder)
    bNotOk = False: Call Sort(SortType)
    Call Refresh
End Sub

'=======================================================
'ne check rien
'=======================================================
Public Sub UnCheckAll()
Dim x As Long
    
    'sélectionne tout
    ReDim bChecked(lListCount)
    
    lCheckCount = 0
    
    'refresh
    Call Refresh
    
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
Dim R As RECT
Dim st As Long
Dim y As Long
Dim z As Long
Dim hRgn As Long
Dim x As Long
Dim hBrush As Long
Dim e As Long
Dim vsEn As Boolean
Static vsMax As Long

    
    If bUnRefreshControl Then Exit Sub
    
    On Error Resume Next
    
    '//on efface
    Call UserControl.Cls

    '//convertit les couleurs
    Call OleTranslateColor(lBackColor, 0, lBackColor)
    Call OleTranslateColor(lBorderColor, 0, lBorderColor)
    
    
    '//on trace chaque élément de la liste
    
    'calcule le nombre d'items qui seront affichés
    x = 0 'contient la hauteur des composants affichés
    z = 0 'contient le nombre d'items à afficher
    For y = lTopIndex To lListCount - 1
        x = x + Col.Item(y).Height
        If x >= Height - 30 Then Exit For
        z = z + 1
    Next y
    
    'limite le Max
    If lListCount <= z + TopIndex Then VS.Max = lListCount - z
    zNumber = z 'sauvegarde le nombre d'Items affichés
    
    If bEnable Then _
        If z < lListCount - 1 Then VS.Enabled = True Else VS.Enabled = False

    'on affiche maintenant chaque controle
    y = 1 'contient la hauteur temporaire
    st = 0
    
    For x = lTopIndex To lTopIndex + z

        'trace le texte
        Call DrawItem(Col.Item(x), y, x)

        'trace l'icone si présente
        If Not (Col.Item(x).Icon = 0) Then
            Call DrawItemIcon(Col.Item(x), y, x)
        End If

        'update la hauteur temporaire
        y = y + Col.Item(x).Height
    Next x


    '//on trace le contour
    If bDisplayBorder Then
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
    End If
    
    
    '//affiche les checkboxes
    If bStyleCheckBox Then Call SplitIMGandShow(z)
    
    'rafraichit le VS si on a changé le Max d'items (permet de changer la
    'hauteur du thumb quand on ajoute des items)
    If vsMax <> VS.Max Then
        vsMax = VS.Max
        Call VS.Refresh
    End If
    
    '//on refresh le control
    Call UserControl.Refresh
    
    bNotOk = True

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
'dessine un item sur le control
'=======================================================
Private Sub DrawItem(Item As vkListItem, lTop As Long, Index As Long)
Dim R As RECT
Dim st As Long
Dim tF As StdFont
Dim o As Long
Dim o2 As Long
Dim F As Long
Dim e As Long
Dim H As Long

    If bVSvisible Then
        'alors on décale le rect pour l'alignement à droite
        o = 19
    Else
        o2 = 17 * Screen.TwipsPerPixelX
    End If
    
    'décalage vers la droite si picture de checkboxes
    If bStyleCheckBox Then
        e = 15
    End If
    If Item.Icon Then
        e = e + Item.pxlIconWidth + 2
    End If
    
    If lFullRowSelect = False Then
        'alors on effectue un décalage si Check
        If bStyleCheckBox Then
            F = 230
        End If
        If Item.Icon Then
            F = F + Screen.TwipsPerPixelX * Item.pxlIconWidth + 80
        End If
    End If
    
    
    'définit la fonte de l'item sur le controle
    Set tF = UserControl.Font
    Set UserControl.Font = Item.Font
    
    'récupère la hauteur du texte à afficher
    H = (Item.Height - TextHeight(Item.Text)) / Screen.TwipsPerPixelY / 2
    
    'définit une zone pour le texte
    Call SetRect(R, 7 + e, 1 + lTop / Screen.TwipsPerPixelY + H, ScaleWidth / Screen.TwipsPerPixelX - 1 - o, _
        1 + lTop / Screen.TwipsPerPixelY + H + TextHeight(Item.Text) / Screen.TwipsPerPixelY)  'lTop + _
        (ScaleHeight - lTop - H / 2) / Screen.TwipsPerPixelY + 1)
    
    'dessine un rectangle (backcolor ou selection) dans cette zone
    If bSelected(Index) = False Then
        'backcolor
        Line (15, lTop + 30)-(Width - 255 - 30 + o2, lTop + Item.Height + 15), Item.BackColor, BF
    Else
        'sélection
        If F Then
            'alors on décale ==> on doit quand même faire le backColor
            Line (15, lTop + 30)-(Width - 255 - 30 + o2, lTop + Item.Height + 15), Item.BackColor, BF
        End If
        
        'fond de la sélection
        Line (15 + F, lTop + 30)-(Width - 255 - 30 + o2, lTop + Item.Height + 15), Item.SelColor, BF
        'bordure de la sélection
        Line (15 + F, lTop + 15)-(Width - 255 - 30 + o2, lTop + 15), Item.BorderSelColor
        Line (Width - 255 - 30 + o2, lTop + 30)-(Width - 255 - 30 + o2, lTop + Item.Height + 15), Item.BorderSelColor
        Line (Width - 255 - 30 + o2, lTop + Item.Height + 15)-(15 + F, lTop + Item.Height + 15), Item.BorderSelColor
        Line (15 + F, lTop + Item.Height + 15)-(15 + F, lTop + 15), Item.BorderSelColor
    End If
        
    
    'prépare l'alignement du texte
    If Item.Alignment = vbLeftJustify Then
        st = DT_LEFT
    ElseIf Item.Alignment = vbCenter Then
        st = DT_CENTER
    Else
        st = DT_RIGHT
    End If
    
    'définit la ForeColor et trace le texte
    Call OleTranslateColor(lForeColor, 0, lForeColor)
    If bEnable Then
        UserControl.ForeColor = Item.ForeColor
    Else
        'couleur de enabled=false
        UserControl.ForeColor = 10070188
    End If
    
    If bUnicode = False Then
        Call DrawText(UserControl.hDC, Item.Text, Len(Item.Text), R, st)
    Else
        Call DrawTextW(UserControl.hDC, StrPtr(Item.Text), Len(Item.Text), R, st)
    End If
    
    Set UserControl.Font = tF 'restaure la fonte d'origine
End Sub

'=======================================================
'dessine l'icone d'un item
'=======================================================
Private Sub DrawItemIcon(Item As vkListItem, lTop As Long, Index As Long)
Dim y As Long
Dim SrcDC As Long
Dim SrcObj As Long
Dim e As Long
Dim pic As StdPicture
    
    'calcule le décalage en haut
    y = 1 + lTop / Screen.TwipsPerPixelY + Item.Height / Screen.TwipsPerPixelY / 2 - Item.pxlIconHeight / 2
    
    'décalage vers la droite si picture de checkboxes
    If bStyleCheckBox Then
        e = 15
    End If
    
    '//si le type est nul, c'est que la picture n'est pas issue d'une
    'icone de fichier par ListType<>NormalList
    If Item.pctType = 0 Then
    
        SrcDC = CreateCompatibleDC(UserControl.hDC)
        SrcObj = SelectObject(SrcDC, Item.Icon)
    
        Call BitBlt(UserControl.hDC, 4 + e, y, Item.pxlIconWidth, _
            Item.pxlIconHeight, SrcDC, 0, 0, SRCCOPY)
    
        Call DeleteDC(SrcDC)
        Call DeleteObject(SrcObj)
    Else
        'alors Icon=1, et la picture est dans tPic()
        
        'si icone ==> DrawIcon, sinon PaintPicture
        'If tPic(Item.Index).Type = vbPicTypeIcon Then
            
        Set pic = GetMyIcon(Item.tagString1)
        
        Call DrawIconEx(hDC, 4 + e, y, pic, Item.pxlIconWidth, _
            Item.pxlIconHeight, 0, 0, DI_NORMAL)
        'Else
            'Call PaintPicture(tPic(Item.Index), 4 + e, y, _
                Item.pxlIconWidth, Item.pxlIconHeight)
        'End If
       
    End If
    
End Sub

Private Sub VS_Change(Value As Currency)
Static lngOldValue
    
    'limite le Max
    If lListCount <= zNumber + TopIndex + 1 Then VS.Max = lListCount - zNumber
    
    lTopIndex = CLng(Value)

    'on en refresh QUE si on a changé de value entre temps
    If lngOldValue <> CLng(Value) Then Call Refresh
    lngOldValue = Value
End Sub

Private Sub VS_Scroll()
    lTopIndex = CLng(VS.Value)
    Call Refresh
End Sub

'=======================================================
'remplit la liste depuis un fichier
'=======================================================
Public Sub FillByFile(ByVal File As String)
Dim lFile As Long
Dim x As Long
Dim s As String
Dim T() As String
    
    On Error Resume Next
    
    'récupère le contenu du fichier
    lFile = FreeFile
    Open File For Binary Access Read As #lFile
    s = Space$(FileLen(File))
    Get #lFile, , s
    Close lFile
    
    'sépare chaque ligne
    ReDim T(0)
    T() = Split(s, vbNewLine, , vbBinaryCompare)
    
    'ajoute tous les items
    bUnRefreshControl = True
    For x = 0 To UBound(T())
        Call Me.AddItem(T(x))
    Next x
    bUnRefreshControl = False
    Call Refresh

End Sub

'=======================================================
'sauve la liste vers un fichier
'=======================================================
Public Sub SaveToFile(ByVal File As String)
Dim lFile As Long
Dim x As Long
Dim s As String
    
    On Error Resume Next
    
    'créé une string depuis les items
    lFile = FreeFile
    Open File For Binary Access Write As #lFile
    For x = 1 To lListCount
        s = Col.Item(x).Text
        If x < lListCount Then
             s = s & vbNewLine
        End If
        Put #lFile, , s
    Next x
    Close lFile

End Sub

'=======================================================
'affiche une des 6 images en la découpant depuis l'image complète
'=======================================================
Private Sub SplitIMGandShow(ByVal z As Long)
Dim hBrush As Long
Dim hRgn As Long
Dim x As Long
Dim y As Single
Dim lIMG As Long
Dim tVal As Boolean
Dim e As Long
    
    'Debug.Print "SplitIMGandShow"
    '0 rien
    '1 survol
    '2 enabled=false
    '3 value enable
    '4 value survol enable
    '5 enable=false OR gray

'    SrcDC = CreateCompatibleDC(UserControl.hdc)
'    SrcObj = SelectObject(SrcDC, CreateCompatibleBitmap(UserControl.hdc, _
'        78, 13))

    'là, on va tracer un rectangle de la couleur BackColor pour effacer les pictures
    'Line (15, 15)-(230, Height - 30), lBackColor, BF
        
    'on découpe l'image correspondant à lIMG depuis Image1 et on blit
    'sur l'usercontrol
    
    If Col.Item(1) Is Nothing Then Exit Sub
    
    For x = lTopIndex To lTopIndex + z
    
        'Top de l'image
        e = y + Col.Item(x).Height / 2 - 100
        
        If bChecked(x) Then
            If MouseItemIndex = x Then
                'checké et survolé
                lIMG = 4
            Else
                'checké sans survol
                lIMG = 3
            End If
        Else
            If MouseItemIndex = x Then
                'pas checké mais survol
                lIMG = 1
            Else
                'pas checké, pas survol
                lIMG = 0
            End If
        End If
        
        'si Enable=false, on change les icones
        If bEnable = False Then
            If bChecked(x) Then
                lIMG = 5
            Else
                lIMG = 2
            End If
        End If
        

        'trace l'image
        Call BitBlt(UserControl.hDC, 2, e / Screen.TwipsPerPixelY, 13, 13, pic(lIMG).hDC, _
              0, 0, SRCCOPY)
        
        'update la hauteur temporaire
        y = y + Col.Item(x).Height
    Next x

    '//on trace le contour
    If bDisplayBorder Then
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
    End If
    
    Call UserControl.Refresh
    'Debug.Print Rnd
    'libère
'    Call DeleteDC(SrcDC)
'    Call DeleteObject(SrcObj)

End Sub

'=======================================================
'trie la collection
'=======================================================
Private Sub Sort(ByVal SortType As SortOrder)
Dim cSort As clsSort
Dim Col2 As clsFastCollection
    
    If SortType = DoNotSort Then Exit Sub
    If bNotOk Or bNotOk2 Or lListCount <= 1 Then Exit Sub
    
    'instancie la classe
    Set cSort = New clsSort
    
    If SortType = Alphabetical Then
        Call cSort.SortList(Col, True)
    Else
        Call cSort.SortList(Col, False)
    End If
    
    'libère
    Set cSort = Nothing
    
    bNotOk = True
    
    'on refresh le controle
    'Call Refresh
End Sub

'=======================================================
'ajoute les fichiers du Path dans la liste
'=======================================================
Private Sub AddFileToList()
Dim s() As String
Dim x As Long
Dim bOk As Boolean
Dim tIt As vkListItem
    
    If bOkToAddFile = False Or bUnRefreshControl Or bNotOk2 Then Exit Sub
    
    On Error GoTo RedimArray
    
    bOk = bUnRefreshControl
    bUnRefreshControl = True
    
    'on clear la sélection
    Call Me.Clear
    
    bUnRefreshControl = bOk
    
    If tListType = SimpleList Then Exit Sub
    
    bUnRefreshControl = True
    
    If tListType = FileListBox Then
        'alors c'est une liste de fichiers
        'récupère tous les fichiers du Path
        Call cFile.EnumFiles(sPath, s(), False, bShowSystemFiles, _
            bShowHiddenFiles, bShowReadOnlyFiles, sPattern)
    ElseIf tListType = FolderListBox Then
        'liste de dossiers
        'énumère tous les dossiers du path
        Call cFile.EnumFolders(sPath, s(), False, bShowSystemFiles, _
            bShowHiddenFiles, bShowReadOnlyFiles)
    Else
        'liste des drives
        'énumères tous les drives
        Call cFile.EnumDrives(s())
    End If
    
    'on vide la collection d'images
    If bDisplayFileIcons Then Call FilePics.Clear
    
    'ajoute tous les items de s()
    For x = 1 To UBound(s())
    
        If bDisplayFileIcons Then
            'on récupère alors l'icone du fichier et on ajoute à la collection
            Set tIt = New vkListItem
            With tIt
                If bDisplayEntirePath Then
                    .Text = s(x)
                Else
                    .Text = GetFileName(s(x))
                End If
                .Icon = 1
                .Font = Ambient.Font
                .Index = x
                .tagString1 = s(x)  'sauve aussi le path entier
                .pctType = 1
                If lIconSize = Size16 Then
                    .pxlIconHeight = 16
                    .pxlIconWidth = 16
                Else
                    .pxlIconHeight = 32
                    .pxlIconWidth = 32
                    .Height = 500
                End If
                .tagString2 = GetFileKey(s(x))
            End With
        
            Call Me.AddItem(, tIt)
        Else
            
            'sans icone
            If bDisplayEntirePath Then
                Call Me.AddItem(s(x))
            Else
                Call Me.AddItem(GetFileName(s(x)))
            End If
            
        End If
        
    Next x
    
    bUnRefreshControl = bOk
    Set tIt = Nothing
    
    'on refresh le controle
    Call Refresh

    Exit Sub
    
RedimArray:
    ReDim tPic(0)
    Resume
End Sub

'=======================================================
'Récupère le nom du fichier depuis le path
'=======================================================
Private Function GetFileName(ByVal Path As String) As String

    Call PathStripPath(Path)
    
    GetFileName = StringWithoutNullChar(Path)
    
End Function

'=======================================================
'Enlève le NullChar
'=======================================================
Private Function StringWithoutNullChar(ByVal strString As String) As String
Dim lIn As Long
    
    lIn = InStr(strString, vbNullChar)
    
    If lIn Then StringWithoutNullChar = Left$(strString, lIn - 1) Else _
        StringWithoutNullChar = strString

End Function

'=======================================================
'Récupère la terminaison d'un fichier
'=======================================================
Private Function GetExt(ByVal File As String) As String
Dim l As Long

    l = InStrRev(File, ".", , vbBinaryCompare)
    If l Then
        GetExt = LCase$(Right$(File, Len(File) - l))
    End If
End Function

'=======================================================
'récupère une Key depuis un fichier
'=======================================================
Private Function GetFileKey(ByVal File As String) As String
Dim s As String

    If tListType = DriveListBox Then
        'on ajoute les drives
        GetFileKey = File
    ElseIf tListType = DriveListBox Then
        'on ajoute des fichiers
        GetFileKey = ":folder"
    Else
        'des fichiers
        s = GetExt(File)
        Select Case s
            Case "exe", "ico", "cur", "ani", "lnk"
                'alors on ajoute comme key le Path
                GetFileKey = File
            Case vbNullString
                'alors on ajoute comme key :no ext
                GetFileKey = ":no ext"
            Case Else
                'alors l'extension
                GetFileKey = s
        End Select
    End If
    
    If lIconSize = Size16 Then
        GetFileKey = GetFileKey & ":1"
    Else
        GetFileKey = GetFileKey & ":3"
    End If
    
End Function

'=======================================================
'Récupère l'icone d'un fichier/dossier/disque
'=======================================================
Private Function GetMyIcon(ByVal s As String) As StdPicture
Dim Key As String

    On Error GoTo NeedToAdd

    'récupère la clé
    Key = GetFileKey(s)
    
    'maintenant qu'on a la key, on essaie d'acceder à l'icone de cette key
    'si erreur ==> on ajoutera
    'si pas d'erreur ==> on ajoute pas
    Set GetMyIcon = picCol.Item(Key)
    
    Exit Function
    
NeedToAdd:

    'on a besoin d'ajouter l'icone
    Call picCol.Add(cFile.GetIcon(s, lIconSize), Key)
    Set GetMyIcon = picCol.Item(Key)
End Function
