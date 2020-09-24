VERSION 5.00
Begin VB.UserControl vkTextBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   MaskColor       =   &H00000000&
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "vkTextBox.ctx":0000
   Begin vkUserContolsXP.vkVScrollPrivate VS 
      Height          =   1815
      Left            =   3960
      TabIndex        =   4
      Top             =   720
      Width           =   220
      _ExtentX        =   397
      _ExtentY        =   3201
   End
   Begin vkUserContolsXP.vkHScrollPrivate HS 
      Height          =   220
      Left            =   1440
      TabIndex        =   3
      Top             =   2760
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   397
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   2
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   2
      Tag             =   "MultiLine textbox with HScroll"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   1
      Tag             =   "MultiLine textbox"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Tag             =   "SingleLine Textbox"
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "vkTextBox"
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
'APIS
'=======================================================
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
Private Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal fnBar As Integer, lPsi As SCROLLINFO) As Boolean

'Private Declare Function SetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal code As Long, ByVal nPos As Long, ByVal fRedraw As Boolean) As Long
'Const WS_EX_STATICEDGE = &H20000
'Const WS_EX_TRANSPARENT = &H20&
'Const WS_CHILD = &H40000000
'Const CW_USEDEFAULT = &H80000000
'Const SW_NORMAL = 1
'Private Type CREATESTRUCT
'    lpCreateParams As Long
'    hInstance As Long
'    hMenu As Long
'    hWndParent As Long
'    cy As Long
'    cx As Long
'    y As Long
'    x As Long
'    style As Long
'    lpszName As String
'    lpszClass As String
'    ExStyle As Long
'End Type
'Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
'Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long


'=======================================================
'CONSTANTES
'=======================================================
Private Const EM_GETSEL = &HB0 ' Gets the starting and ending character positions of the current selection.
Private Const EM_SETSEL = &HB1 ' Positions the cursor at a specific character location, or selects a range of characters.
Private Const EM_SCROLL = &HB5 ' Scrolls the text vertically in a multiline edit control.
Private Const EM_LINESCROLL = &HB6 ' Scrolls the text vertically or horizontally in a multiline edit control. Use this message to scroll to a specific line or character position.
Private Const EM_SCROLLCARET = &HB7 ' Scrolls the caret into view.
Private Const EM_GETLINECOUNT = &HBA ' Returns the number of lines in a multiline edit control. If no text is in the control, the return value is 1.
Private Const EM_LINEINDEX = &HBB ' Retrieves the number of characters from the beginning of the edit control to the beginning of the specified line.
Private Const EM_LINELENGTH = &HC1 ' Retrieves the amount of characters in a specified line in the edit control. Use the EM_LINEINDEX message to retrieve a character index for a given line number within a multiline edit control.
Private Const EM_GETLINE = &HC4 ' Copies a specific line of text and places it in the specified buffer. Note: The copied line does not contain a terminating null character. The return value is the number of characters copied.
Private Const EM_CANUNDO = &HC6 ' Returns whether an edit-control operation can be undone; that is, whether the control can respond to the EM_UNDO message.
Private Const EM_UNDO = &HC7 ' An application sends this message to undo the last edit control operation.
Private Const EM_LINEFROMCHAR = &HC9 ' Retrieves the line number of the line that contains the specified character index. A character index is specified as the number of characters from the beginning of the control.
Private Const EM_EMPTYUNDOBUFFER = &HCD ' Resets the undo flag. The undo flag is set whenever an operation within the edit control can be undone.
Private Const EM_GETFIRSTVISIBLELINE = &HCE ' Returns the zero-based index of the uppermost visible line in a multiline edit control, or the zero-based index of the first visible character for a single-line edit control.
Private Const EM_SETMARGINS = &HD3 ' Sets the widths of the left and right margins. The message redraws the control to reflect the new margins.
Private Const EM_GETMARGINS = &HD4 ' Retrieves the widths of the left and right margins. Returns the width of the left margin in the low-order word, and the width of the right margin in the high-order word.
Private Const EM_POSFROMCHAR = &HD6 ' Retrieves the coordinates of the specified zero-based character index in an edit control.
Private Const EM_CHARFROMPOS = &HD7 ' Retrieves the character index and line index of the character nearest a specified point in the client area of an edit control.
Private Const EC_LEFTMARGIN = &H1   ' Right margin
Private Const EC_RIGHTMARGIN = &H2  'Left margin
Private Const WM_COPY               As Long = &H301
Private Const WM_CUT                As Long = &H300
Private Const WM_PASTE              As Long = &H302
Private Const SB_LINEDOWN           As Long = 1
Private Const SB_LINELEFT           As Long = 0
Private Const SB_LINERIGHT          As Long = 1
Private Const SB_LINEUP             As Long = 0
Private Const SB_PAGEDOWN           As Long = 3
Private Const SB_PAGELEFT           As Long = 2
Private Const SB_PAGERIGHT          As Long = 3
Private Const SB_PAGEUP             As Long = 2
Private Const SB_HORZ               As Long = 0
Private Const SB_VERT               As Long = 1
Private Const SM_CYHSCROLL          As Long = 3
Private Const SIF_RANGE             As Long = &H1
Private Const SIF_PAGE              As Long = &H2
Private Const SIF_POS               As Long = &H4
Private Const SIF_DISABLENOSCROLL   As Long = &H8
Private Const SIF_TRACKPOS          As Long = &H10
Private Const SIF_ALL               As Long = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)


'=======================================================
'VARIABLES PRIVEES
'=======================================================
Private mAsm(63) As Byte    'contient le code ASM
Private OldProc As Long     'adresse de l'ancienne window proc
Private objHwnd As Long     'handle de l'objet concerné
Private ET As TRACKMOUSEEVENTTYPE   'type pour le mouse_hover et le mouse_leave
Private IsMouseIn As Boolean    'si la souris est dans le controle

Private bNotOk As Boolean
Private bNotOk2 As Boolean
Private bUnRefreshControl As Boolean

Private bDisplayBorder As Boolean
Private lBorderColor As OLE_COLOR
Private bEnable As Boolean
Private lForeColor As OLE_COLOR
Private tScrollBars As ScrollBarConstants
Private sText As String
Private lSelStart As Long
Private lSelLength As Long
Private sLegendText As String
Private tLegendAlignmentY As AlignmentConstants
Private tLegendAlignmentX As AlignmentConstants
Private bLegendAutoSize As Boolean
Private lLegendBackColor1 As OLE_COLOR
Private lLegendBackColor2 As OLE_COLOR
Private tLegendGradient As GradientConstants
Private lLegendForeColor As OLE_COLOR
Private tLegendType As LegendTypeConstants
Private lLegendWidth As Long
Private tLegendOrientation As OrientationsConstants
Private lLeftMargin As Long
Private lRightMargin As Long
Private lCurrentTextBox As Long
Private sAuthorizedChars As String
Private lCharH As Long
Private bNot3 As Long
Private bDamn As Boolean
Private bDoNotSync As Boolean
Private bOk4 As Boolean
Private bUnicode As Boolean
'Private tScroll As New vkListBoxVScroll


'=======================================================
'EVENTS
'=======================================================
Public Event Change()
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event LegendKeyDown(KeyCode As Integer, Shift As Integer)
Public Event LegendKeyPress(KeyAscii As Integer)
Public Event LegendKeyUp(KeyCode As Integer, Shift As Integer)
Public Event LegendMouseHover()
Public Event LegendMouseLeave()
Public Event LegendMouseWheel(Sens As Wheel_Sens)
Public Event LegendMouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
Public Event LegendMouseUp(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
Public Event LegendMouseDblClick(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
Public Event LegendMouseMove(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
Public Event VScrolled(Value As Currency)
Public Event HScrolled(Value As Currency)



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
Static FirstV As Long
Static HVal As Long
Dim F As Long
    
    Select Case uMsg
        
        Case 307, 32
            'on a changé le FirstVisible
            VS.BlockVS = True
            
            'bDamn est là pour éviter que le subclassing empeche de
            'changer VS.Value sur le controle
            
            '//on calcule le Value à afficher
            'on ne change QUE si le FirstVisible a changé
            F = FirstVisibleLine
            If FirstV <> F Then
                'calcule le nombre de lignes visibles
                z = Int(Text1(1).Height / TextHeight("A"))
                'change le Value
                If bDamn = False Then VS.Value = FirstVisibleLine + 1 Else bDamn = False
        
                'If bDamn = False Then VS.Value = LineIndex + 1
                    
            End If
            FirstV = F
            'met à jour les numéros de ligne
            bNot3 = 0: Call ChangeNum
            
            '//synchronise les Scrollbars
'            If HVal <> HS.Value Then
                Call SyncHS
'            End If
'            HVal = HS.Value
           
        Case WM_LBUTTONDBLCLK
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
            RaiseEvent LegendMouseDblClick(vbLeftButton, iShift, iControl, x, y)
        Case WM_LBUTTONDOWN
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
                RaiseEvent LegendMouseDown(vbLeftButton, iShift, iControl, x, y)
        Case WM_LBUTTONUP
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
                RaiseEvent LegendMouseUp(vbLeftButton, iShift, iControl, x, y)
        Case WM_MBUTTONDBLCLK
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
                RaiseEvent LegendMouseDblClick(vbMiddleButton, iShift, iControl, x, y)
        Case WM_MBUTTONDOWN
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
                RaiseEvent LegendMouseDown(vbMiddleButton, iShift, iControl, x, y)
        Case WM_MBUTTONUP
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
                RaiseEvent LegendMouseUp(vbMiddleButton, iShift, iControl, x, y)
        Case WM_MOUSEHOVER
            If IsMouseIn = False Then
                RaiseEvent LegendMouseHover
                IsMouseIn = True
            End If
        Case WM_MOUSELEAVE
            RaiseEvent LegendMouseLeave
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
                RaiseEvent LegendMouseMove(z, iShift, iControl, x, y)
        Case WM_RBUTTONDBLCLK
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
                RaiseEvent LegendMouseDblClick(vbRightButton, iShift, iControl, x, y)
        Case WM_RBUTTONDOWN
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
                RaiseEvent LegendMouseDown(vbRightButton, iShift, iControl, x, y)
        Case WM_RBUTTONUP
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * Screen.TwipsPerPixelX
                y = HiWord(lParam) * Screen.TwipsPerPixelY
                
                RaiseEvent LegendMouseUp(vbRightButton, iShift, iControl, x, y)
        Case WM_MOUSEWHEEL
            If wParam < 0 Then
                RaiseEvent LegendMouseWheel(WHEEL_DOWN)
            Else
                RaiseEvent LegendMouseWheel(WHEEL_UP)
            End If
        Case WM_PAINT
            bNotOk = True  'évite le clignotement lors du survol de la souris
    End Select
    
    'appel de la routine standard pour les autres messages
    WindowProc = CallWindowProc(OldProc, hWnd, uMsg, wParam, lParam)
    
End Function

'=======================================================
'SIMPLE EVENTS
'=======================================================
Private Sub Text1_Click(Index As Integer): RaiseEvent Click: End Sub
Private Sub Text1_DblClick(Index As Integer): RaiseEvent DblClick: End Sub
Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer): RaiseEvent KeyDown(KeyCode, Shift): End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer): RaiseEvent KeyPress(KeyAscii): End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer): RaiseEvent KeyUp(KeyCode, Shift): End Sub
Private Sub Text1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single): RaiseEvent MouseDown(Button, Shift, x, y): End Sub
Private Sub Text1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single): RaiseEvent MouseMove(Button, Shift, x, y): End Sub
Private Sub Text1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single): RaiseEvent MouseUp(Button, Shift, x, y): End Sub

Private Sub Text1_Change(Index As Integer)
Dim zNumber As Long
Dim o As Boolean
Dim lPsi As SCROLLINFO
    
    sText = Text1(Index).Text
    
    'nombre d'Items visibles
    zNumber = Int(Text1(1).Height / TextHeight("A"))
    
    '//on enable ou pas le HS
    lPsi.cbSize = Len(lPsi)
    lPsi.fMask = SIF_PAGE
    Call GetScrollInfo(Text1(2).hWnd, SB_HORZ, lPsi)
    o = HS.Enabled
    HS.Enabled = (lPsi.nPage <> 1)
    
    'change le max des deux scrolls
    VS.Max = LineCount - zNumber + 1
    
    '//on synchronise les HSCROLLs
    Call SyncHS

    '//on change la propriété Enable du VS et du HS si besoin
    VS.Enabled = (LineCount > zNumber)
    
    'Call Refresh 'pour changer les lignes affichées
    RaiseEvent Change
    
    Call ChangeNum(True)
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
    
    'initialise leS VS
    bUnRefreshControl = True
    With VS
        .UnRefreshControl = True
        .Max = 2
        .Min = 1
        .Value = 1
        .Visible = False
        .Enabled = False
        .UnRefreshControl = False
        'Call .Refresh
    End With
    With HS
        .UnRefreshControl = True
        .Max = 2
        .Min = 0
        .Value = 1
        .Visible = False
        .Enabled = False
        .UnRefreshControl = False
        .SmallChange = 1
        .LargeChange = 3
        .WheelChange = 2
        'Call .Refresh
    End With
    bUnRefreshControl = False
End Sub

Private Sub UserControl_InitProperties()
    'valeurs par défaut
    bNotOk2 = True
    With Me
        .Text = vbNullString
        .ForeColor = vbBlack
        .BackColor = vbWhite
        .LegendFont = Ambient.Font
        .Font = Ambient.Font
        .Enabled = True
        .Alignment = vbLeftJustify
        .UnRefreshControl = False
        .BorderColor = &HFF8080
        .Locked = False
        .MaxLength = 0
        '.MultiLine = False
        .PassWordChar = vbNullString
        .DisplayBorder = True
        .ScrollBars = vbSBNone
        .LegendText = vbNullString
        .LegendAlignmentX = vbCenter
        .LegendAlignmentY = vbLeftJustify
        .LegendAutosize = False
        .LegendBackColor2 = &HEDD0BF
        .LegendBackColor1 = &HEDD0BF
        .LegendGradient = None
        .LegendForeColor = &H8000000D
        .LegendType = NoLegend
        .LegendWidth = 250
        .LegendOrientation = VerticalLegend
        .AuthorizedChars = vbNullString
        .UseUnicode = False
    End With
    bNotOk2 = False
    Call UserControl_Paint  'refresh
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent LegendKeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent LegendKeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent LegendKeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_Terminate()
    'vire le subclassing
    If OldProc Then Call SetWindowLong(UserControl.hWnd, GWL_WNDPROC, OldProc)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Text", Me.Text, vbNullString)
        Call .WriteProperty("ForeColor", Me.ForeColor, vbBlack)
        Call .WriteProperty("BackColor", Me.BackColor, vbWhite)
        Call .WriteProperty("LegendFont", Me.LegendFont, Ambient.Font)
        Call .WriteProperty("Font", Me.Font, Ambient.Font)
        Call .WriteProperty("Enabled", Me.Enabled, True)
        Call .WriteProperty("Alignment", Me.Alignment, vbLeftJustify)
        Call .WriteProperty("UnRefreshControl", Me.UnRefreshControl, False)
        Call .WriteProperty("BorderColor", Me.BorderColor, &HFF8080)
        Call .WriteProperty("Locked", Me.Locked, False)
        Call .WriteProperty("MaxLength", Me.MaxLength, 0)
        Call .WriteProperty("MultiLine", Me.MultiLine, False)
        Call .WriteProperty("PassWordChar", Me.PassWordChar, vbNullString)
        Call .WriteProperty("DisplayBorder", Me.DisplayBorder, True)
        Call .WriteProperty("ScrollBars", Me.ScrollBars, vbSBNone)
        Call .WriteProperty("LegendText", Me.LegendText, vbNullString)
        Call .WriteProperty("LegendAlignmentY", Me.LegendAlignmentY, vbLeftJustify)
        Call .WriteProperty("LegendAlignmentX", Me.LegendAlignmentX, vbCenter)
        Call .WriteProperty("LegendAutosize", Me.LegendAutosize, False)
        Call .WriteProperty("LegendBackColor1", Me.LegendBackColor1, &HEDD0BF)
        Call .WriteProperty("LegendBackColor2", Me.LegendBackColor2, &HEDD0BF)
        Call .WriteProperty("LegendGradient", Me.LegendGradient, None)
        Call .WriteProperty("LegendForeColor", Me.LegendForeColor, &H8000000D)
        Call .WriteProperty("LegendType", Me.LegendType, NoLegend)
        Call .WriteProperty("LegendWidth", Me.LegendWidth, 250)
        Call .WriteProperty("LegendOrientation", Me.LegendOrientation, VerticalLegend)
        Call .WriteProperty("AuthorizedChars", Me.AuthorizedChars, vbNullString)
        Call .WriteProperty("UseUnicode", Me.UseUnicode, False)
    End With
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    bNotOk2 = True
    With PropBag
        Set Me.Font = .ReadProperty("Font", Ambient.Font)
        Set Me.LegendFont = .ReadProperty("LegendFont", Ambient.Font)
        Me.Text = .ReadProperty("Text", vbNullString)
        Me.ForeColor = .ReadProperty("ForeColor", vbBlack)
        Me.BackColor = .ReadProperty("BackColor", vbWhite)
        Me.Enabled = .ReadProperty("Enabled", True)
        Me.Alignment = .ReadProperty("Alignment", vbLeftJustify)
        Me.UnRefreshControl = .ReadProperty("UnRefreshControl", False)
        Me.BorderColor = .ReadProperty("BorderColor", &HFF8080)
        Me.Locked = .ReadProperty("Locked", False)
        Me.MaxLength = .ReadProperty("MaxLength", 0)
        Me.MultiLine = .ReadProperty("MultiLine", False)
        Me.PassWordChar = .ReadProperty("PassWordChar", vbNullString)
        Me.DisplayBorder = .ReadProperty("DisplayBorder", True)
        Me.ScrollBars = .ReadProperty("ScrollBars", vbSBNone)
        Me.LegendText = .ReadProperty("LegendText", vbNullString)
        Me.LegendAlignmentY = .ReadProperty("LegendAlignmentY", vbLeftJustify)
        Me.LegendAlignmentX = .ReadProperty("LegendAlignmentX", vbCenter)
        Me.LegendAutosize = .ReadProperty("LegendAutosize", False)
        Me.LegendBackColor1 = .ReadProperty("LegendBackColor1", &HEDD0BF)
        Me.LegendBackColor2 = .ReadProperty("LegendBackColor2", &HEDD0BF)
        Me.LegendGradient = .ReadProperty("LegendGradient", None)
        Me.LegendForeColor = .ReadProperty("LegendForeColor", &H8000000D)
        Me.LegendType = .ReadProperty("LegendType", NoLegend)
        Me.LegendWidth = .ReadProperty("LegendWidth", 250)
        Me.LegendOrientation = .ReadProperty("LegendOrientation", VerticalLegend)
        Me.AuthorizedChars = .ReadProperty("AuthorizedChars", vbNullString)
        Me.UseUnicode = .ReadProperty("UseUnicode", False)
       ' Me.VScroll = .ReadProperty("VScroll", tScroll)
    End With
    bNotOk2 = False
    'Call UserControl_Paint  'refresh
    Call RefreshPrivate
    
    'le bon endroit pour lancer le subclassing
    Call LaunchKeyMouseEvents
    
    If Ambient.UserMode Then
        Call VS.LaunchKeyMouseEvents    'subclasse également le VS
        Call HS.LaunchKeyMouseEvents    'et le HS
    End If
    
End Sub
Private Sub UserControl_Resize()
Dim V As Long
Dim H As Long
Dim hRgn As Long

    'on redimensionne les textboxes en fonction de la taille du controle et
    'de la Legend

    If tScrollBars = vbBoth Or tScrollBars = vbVertical Then
        'alors Vscroll ==> on décale
        V = VS.Width
    End If
    If tScrollBars = vbBoth Or tScrollBars = vbHorizontal Then
        'alors Vscroll ==> on décale
        H = HS.Height
    End If
    
    'on positionne la VScroll
    With VS
        .Left = Width - 30 - V
        .Top = 30
        .Height = Height - 60 - H
    End With
    'on positionne la HScroll
    With HS
        .Left = lLegendWidth + 30
        .Top = Height - H - 30
        .Width = Width - .Left - 30 - V
    End With
    
    If tLegendType <> NoLegend Then
        'alors Legend ==> on décale
        With Text1(0)
            .Left = lLegendWidth + 45
            .Top = 30
            .Width = Width - 90 - lLegendWidth - V
            .Height = Height - 45 - H
        End With
        With Text1(1)
            .Left = Text1(0).Left
            .Top = Text1(0).Top
            .Width = Text1(0).Width
            .Height = Text1(0).Height
        End With
        With Text1(2)
            .Left = Text1(0).Left
            .Top = Text1(0).Top
            .Width = Text1(0).Width
            .Height = Text1(0).Height + 14 * GetSystemMetrics(SM_CYHSCROLL)
        End With
    Else
        With Text1(0)
            .Left = 45
            .Top = 30
            .Width = Width - 90 - V
            .Height = Height - 45 - H
        End With
        With Text1(1)
            .Left = Text1(0).Left
            .Top = Text1(0).Top
            .Width = Text1(0).Width
            .Height = Text1(0).Height
        End With
        With Text1(2)
            .Left = Text1(0).Left
            .Top = Text1(0).Top
            .Width = Text1(0).Width
            .Height = Text1(0).Height + 14 * GetSystemMetrics(SM_CYHSCROLL)
        End With
    End If

    If MultiLine Then
        lCurrentTextBox = 1
        Text1(1).Visible = True
        Text1(0).Visible = False
    Else
        lCurrentTextBox = 0
        Text1(0).Visible = True
        Text1(1).Visible = False
    End If
    
    'maintenant on "cache" la HScroll classique ;)
    'c'est le seule moyen de la garder (pour récupérer sa property Max)
    'tout en la masquant !
    hRgn = CreateRectRgn(0, 0, Text1(2).Width / Screen.TwipsPerPixelX, _
            Text1(2).Height / Screen.TwipsPerPixelY - GetSystemMetrics(SM_CYHSCROLL))
    Call SetWindowRgn(Text1(2).hWnd, hRgn, True)
    Call DeleteObject(hRgn)
    
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
Public Property Get Text() As String: Text = sText: End Property
Public Property Let Text(Text As String): sText = Text: bNotOk = False: UserControl_Paint: bNotOk = True: Text1(1).Text = sText: Text1(0).Text = sText: Text1(2).Text = sText: End Property
Public Property Get ForeColor() As OLE_COLOR: ForeColor = lForeColor: End Property
Public Property Let ForeColor(ForeColor As OLE_COLOR): lForeColor = ForeColor: UserControl.ForeColor = ForeColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get BackColor() As OLE_COLOR: BackColor = Text1(0).BackColor: End Property
Public Property Let BackColor(BackColor As OLE_COLOR): Text1(0).BackColor = BackColor: Text1(1).BackColor = BackColor: Text1(2).BackColor = BackColor: UserControl.BackColor = BackColor: End Property
Public Property Get LegendFont() As StdFont: Set LegendFont = UserControl.Font: End Property
Public Property Set LegendFont(LegendFont As StdFont): Set UserControl.Font = LegendFont: bNotOk = False: UserControl_Paint: End Property
Public Property Get Font() As StdFont: Set Font = Text1(0).Font: End Property
Public Property Set Font(Font As StdFont)
'on récupère la hauteur de la police
Dim oldFont As StdFont
    'ancienne font (celle de Legend)
    Set oldFont = UserControl.Font
    'nouvelle font qu'on applique sur l'usercontrol pour bénéficier d'un DC
    Set UserControl.Font = Font
    'récupère la hauteur
    lCharH = GetCharHeight
    'remet l'ancienne font sur le usercontrol
    Set UserControl.Font = oldFont
    'libère
    Set oldFont = Nothing
Set Text1(0).Font = Font: Set Text1(1).Font = Font: Set Text1(2).Font = Font
End Property
Public Property Get Enabled() As Boolean: Enabled = bEnable: End Property
Public Property Let Enabled(Enabled As Boolean): bEnable = Enabled: Text1(0).Enabled = Enabled: Text1(1).Enabled = Enabled: Text1(2).Enabled = Enabled: bNotOk = False: UserControl_Paint: End Property
Public Property Get Alignment() As AlignmentConstants: Alignment = Text1(0).Alignment: End Property
Public Property Let Alignment(Alignment As AlignmentConstants): Text1(0).Alignment = Alignment: Text1(1).Alignment = Alignment: Text1(2).Alignment = Alignment: End Property
Public Property Get UnRefreshControl() As Boolean: UnRefreshControl = bUnRefreshControl: End Property
Public Property Let UnRefreshControl(UnRefreshControl As Boolean): bUnRefreshControl = UnRefreshControl: End Property
Public Property Get BorderColor() As OLE_COLOR: BorderColor = lBorderColor: End Property
Public Property Let BorderColor(BorderColor As OLE_COLOR): lBorderColor = BorderColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get Locked() As Boolean: Locked = Text1(0).Locked: End Property
Public Property Let Locked(Locked As Boolean): Text1(0).Locked = Locked: Text1(1).Locked = Locked: Text1(2).Locked = Locked: End Property
Public Property Get MaxLength() As Long: MaxLength = Text1(0).MaxLength: End Property
Public Property Let MaxLength(MaxLength As Long)
MaxLength = Abs(MaxLength): Text1(0).MaxLength = MaxLength: Text1(1).MaxLength = MaxLength: Text1(2).MaxLength = MaxLength: End Property
Public Property Get MultiLine() As Boolean
    If lCurrentTextBox = 0 Then
        MultiLine = False
    Else
        MultiLine = True
    End If
End Property
Public Property Let MultiLine(MultiLine As Boolean)
    If MultiLine And tScrollBars <> vbHorizontal Then
        lCurrentTextBox = 1
        Text1(1).Text = sText
        Text1(1).Visible = True
        Text1(0).Visible = False
        Text1(2).Visible = False
    ElseIf Not (MultiLine) Then
        Text1(0).Text = sText
        lCurrentTextBox = 0
        Text1(0).Visible = True
        Text1(1).Visible = False
        Text1(2).Visible = False
    Else
        Text1(2).Text = sText
        lCurrentTextBox = 2
        Text1(0).Visible = False
        Text1(1).Visible = False
        Text1(2).Visible = True
    End If
End Property
Public Property Get PassWordChar() As String: PassWordChar = Text1(0).PassWordChar: End Property
Public Property Let PassWordChar(PassWordChar As String)
If Len(PassWordChar) > 1 Then Exit Property
Text1(0).PassWordChar = PassWordChar: Text1(1).PassWordChar = PassWordChar: Text1(2).PassWordChar = PassWordChar
End Property
Public Property Get DisplayBorder() As Boolean: DisplayBorder = bDisplayBorder: End Property
Public Property Let DisplayBorder(DisplayBorder As Boolean): bDisplayBorder = DisplayBorder: bNotOk = False: UserControl_Paint: End Property
Public Property Get ScrollBars() As ScrollBarConstants: ScrollBars = tScrollBars: End Property
Public Property Let ScrollBars(ScrollBars As ScrollBarConstants)
tScrollBars = ScrollBars
If tScrollBars = vbBoth Or tScrollBars = vbVertical Then VS.Visible = True Else VS.Visible = False
If tScrollBars = vbBoth Or tScrollBars = vbHorizontal Then HS.Visible = True Else HS.Visible = False
If MultiLine And tScrollBars = vbVertical Then
    lCurrentTextBox = 1
    Text1(1).Text = sText
    Text1(1).Visible = True
    Text1(0).Visible = False
    Text1(2).Visible = False
ElseIf Not (MultiLine) Then
    Text1(0).Text = sText
    lCurrentTextBox = 0
    Text1(0).Visible = True
    Text1(1).Visible = False
    Text1(2).Visible = False
Else
    Text1(2).Text = sText
    lCurrentTextBox = 2
    Text1(0).Visible = False
    Text1(1).Visible = False
    Text1(2).Visible = True
End If
UserControl_Resize
End Property
Public Property Get SelStart() As Long: SelStart = Text1(0).SelStart: End Property
Public Property Let SelStart(SelStart As Long): Text1(0).SelStart = SelStart: Text1(1).SelStart = SelStart: Text1(2).SelStart = SelStart: End Property
Public Property Get SelLength() As Long: SelLength = Text1(0).SelLength: End Property
Public Property Let SelLength(SelLength As Long): Text1(0).SelLength = SelLength: Text1(1).SelLength = SelLength: Text1(2).SelLength = SelLength: End Property
Public Property Get SelText() As String: SelText = Text1(0).SelText: End Property
Public Property Let SelText(SelText As String): Text1(0).SelText = SelText: Text1(1).SelText = SelText: Text1(2).SelText = SelText: End Property
Public Property Get LegendText() As String: LegendText = sLegendText: End Property
Public Property Let LegendText(LegendText As String)
If LegendAutosize And tLegendOrientation = HorizontalLegend Then lLegendWidth = TextWidth(LegendText): UserControl_Resize
sLegendText = LegendText: bNotOk = False: UserControl_Paint: bNotOk = True
End Property
Public Property Get LegendAlignmentX() As AlignmentConstants: LegendAlignmentX = tLegendAlignmentX: End Property
Public Property Let LegendAlignmentX(LegendAlignmentX As AlignmentConstants): tLegendAlignmentX = LegendAlignmentX: bNotOk = False: UserControl_Paint: bNotOk = True: End Property
Public Property Get LegendAlignmentY() As AlignmentConstants: LegendAlignmentY = tLegendAlignmentY: End Property
Public Property Let LegendAlignmentY(LegendAlignmentY As AlignmentConstants): tLegendAlignmentY = LegendAlignmentY: bNotOk = False: UserControl_Paint: bNotOk = True: End Property
Public Property Get LegendAutosize() As Boolean: LegendAutosize = bLegendAutoSize: End Property
Public Property Let LegendAutosize(LegendAutosize As Boolean)
If LegendAutosize And tLegendOrientation = HorizontalLegend And tLegendType = SimpleTextLegend Then lLegendWidth = TextWidth(LegendText)
UserControl_Resize
bLegendAutoSize = LegendAutosize: bNotOk = False: UserControl_Paint: bNotOk = True
End Property
Public Property Get LegendBackColor1() As OLE_COLOR: LegendBackColor1 = lLegendBackColor1: End Property
Public Property Let LegendBackColor1(LegendBackColor1 As OLE_COLOR): lLegendBackColor1 = LegendBackColor1: bNotOk = False: UserControl_Paint: bNotOk = True: End Property
Public Property Get LegendBackColor2() As OLE_COLOR: LegendBackColor2 = lLegendBackColor2: End Property
Public Property Let LegendBackColor2(LegendBackColor2 As OLE_COLOR): lLegendBackColor2 = LegendBackColor2: bNotOk = False: UserControl_Paint: bNotOk = True: End Property
Public Property Get LegendGradient() As GradientConstants: LegendGradient = tLegendGradient: End Property
Public Property Let LegendGradient(LegendGradient As GradientConstants): tLegendGradient = LegendGradient: bNotOk = False: UserControl_Paint: bNotOk = True: End Property
Public Property Get LegendForeColor() As OLE_COLOR: LegendForeColor = lLegendForeColor: End Property
Public Property Let LegendForeColor(LegendForeColor As OLE_COLOR): lLegendForeColor = LegendForeColor: bNotOk = False: UserControl_Paint: bNotOk = True: End Property
Public Property Get LegendType() As LegendTypeConstants: LegendType = tLegendType: End Property
Public Property Let LegendType(LegendType As LegendTypeConstants): tLegendType = LegendType: bNotOk = False: UserControl_Paint: bNotOk = True: UserControl_Resize: End Property
Public Property Get LegendWidth() As Long: LegendWidth = lLegendWidth: End Property
Public Property Let LegendWidth(LegendWidth As Long)
lLegendWidth = LegendWidth
If LegendAutosize And tLegendOrientation = HorizontalLegend And tLegendType = SimpleTextLegend Then lLegendWidth = TextWidth(LegendText)
UserControl_Resize
End Property
Public Property Get LegendOrientation() As OrientationsConstants: LegendOrientation = tLegendOrientation: End Property
Public Property Let LegendOrientation(LegendOrientation As OrientationsConstants)
tLegendOrientation = LegendOrientation
If LegendAutosize And tLegendOrientation = HorizontalLegend And tLegendType = SimpleTextLegend Then lLegendWidth = TextWidth(LegendText)
UserControl_Resize
End Property
Public Property Get LeftMargin() As Long: LeftMargin = lLeftMargin: End Property
Public Property Let LeftMargin(LeftMargin As Long): lLeftMargin = LeftMargin: Call PostMessage(Text1(0).hWnd, EM_SETMARGINS, EC_LEFTMARGIN, LeftMargin): Call PostMessage(Text1(1).hWnd, EM_SETMARGINS, EC_LEFTMARGIN, LeftMargin): Call PostMessage(Text1(2).hWnd, EM_SETMARGINS, EC_LEFTMARGIN, LeftMargin): End Property
Public Property Get RightMargin() As Long: RightMargin = lRightMargin: End Property
Public Property Let RightMargin(RightMargin As Long): lRightMargin = RightMargin: Call PostMessage(Text1(0).hWnd, EM_SETMARGINS, EC_RIGHTMARGIN, RightMargin): Call PostMessage(Text1(1).hWnd, EM_SETMARGINS, EC_RIGHTMARGIN, RightMargin): Call PostMessage(Text1(2).hWnd, EM_SETMARGINS, EC_RIGHTMARGIN, RightMargin): End Property
Public Property Get FirstVisibleLine() As Long: FirstVisibleLine = SendMessage(Text1(lCurrentTextBox).hWnd, EM_GETFIRSTVISIBLELINE, 0, 0): End Property
Public Property Get LineCount() As Long
LineCount = SendMessage(Text1(lCurrentTextBox).hWnd, EM_GETLINECOUNT, 0, 0)
End Property
Public Property Get FirstChar(Optional ByVal LineIndex As Long = -1)
If LineIndex = -1 Then LineIndex = Me.LineIndex(Text1(lCurrentTextBox).SelStart)
FirstChar = SendMessage(Text1(lCurrentTextBox).hWnd, EM_LINEINDEX, LineIndex, 0)
End Property
Public Property Get LineLength(Optional ByVal LineIndex As Long = -1, Optional ByVal CharPos = -1)
If LineIndex <> -1 Then CharPos = FirstChar(LineIndex)
If CharPos = -1 Then CharPos = Text1(lCurrentTextBox).SelStart
LineLength = SendMessage(Text1(lCurrentTextBox).hWnd, EM_LINELENGTH, CharPos, 0)
End Property
Public Property Get LineIndex(Optional ByVal CharPos = -1) As Long
If CharPos = -1 Then CharPos = Text1(lCurrentTextBox).SelStart
LineIndex = 1 + SendMessage(Text1(lCurrentTextBox).hWnd, EM_LINEFROMCHAR, CharPos, 0)
End Property
Public Property Get AuthorizedChars() As String: AuthorizedChars = sAuthorizedChars: End Property
Public Property Let AuthorizedChars(AuthorizedChars As String): sAuthorizedChars = AuthorizedChars: Call CheckChars: End Property
Public Property Let LineNumber(LineIndex As Long)
    LineIndex = LineIndex - 1
    If LineIndex > LineCount Or LineIndex < 0 Then Exit Property
    Text1(lCurrentTextBox).SelStart = SendMessage(Text1(lCurrentTextBox).hWnd, EM_LINEINDEX, LineIndex, 0)
    Call Me.EnsureCaretVisibility
End Property
Public Property Let FirstVisibleLine(NewLine As Long)
Call SendMessage(Text1(lCurrentTextBox).hWnd, EM_LINESCROLL, 0, _
        CLng(NewLine - 1 - FirstVisibleLine))
End Property
'Public Property Get VSValue() As Long: VSValue = VS.Value: End Property
'Public Property Get VSMax() As Long: VSMax = VS.Max: End Property
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
        .Width = VScroll.Width
        .Left = Width - .Width
        .UnRefreshControl = False
    End With
    Call VS.Refresh
    bNotOk = False: Call UserControl_Paint
End Property
Public Property Get UseUnicode() As Boolean: UseUnicode = bUnicode: End Property
Public Property Let UseUnicode(UseUnicode As Boolean): bUnicode = UseUnicode: bNotOk = False: UserControl_Paint: bNotOk = True: End Property


Private Sub UserControl_Paint()

    If bNotOk Or bNotOk2 Then Exit Sub            'pas prêt à peindre
    
    bNot3 = 1
    Call RefreshPrivate    'on refresh
End Sub




'=======================================================
'PRIVATE SUBS
'=======================================================
'=======================================================
'vire tous les caractères non conformes
'=======================================================
Private Sub CheckChars()
Dim x As Long
Dim z As Long
Dim y As Long
Dim s() As String

    If sAuthorizedChars = vbNullString Then Exit Sub
    
    'divise le texte en lignes
    ReDim s(0)
    s() = Split(sText, vbNewLine, , vbBinaryCompare)
    
    'on vire tous les caractères que l'on refuse
    For x = 0 To UBound(s())
        z = Len(s(x))
        For y = z To 1 Step -1
            If InStr(1, sAuthorizedChars, Mid$(s(x), y, 1), vbBinaryCompare) = 0 Then
                'alors on vire ce char
                s(x) = Left$(s(x), y - 1) & Right$(s(x), Len(s(x)) - y)
            End If
        Next y
                
    Next x
    
    'on recréé notre string
    sText = vbNullString
    For x = 0 To UBound(s())
        sText = sText & s(x)
        If x < UBound(s()) Then sText = sText & vbNewLine
    Next x
    
    Text1(lCurrentTextBox).Text = sText
End Sub

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
'récupère la hauteur d'un caractère du DC du controle
'=======================================================
Private Function GetCharHeight() As Long
Dim Res As Long
    Res = GetTabbedTextExtent(UserControl.hDC, "A", 1, 0, 0)
    GetCharHeight = (Res And &HFFFF0000) \ &H10000
End Function

'=======================================================
'PUBLIC FUNCTIONS
'=======================================================
'=======================================================
'effectue un Undo
'=======================================================
Public Sub Undo()
    Call PostMessage(Text1(0).hWnd, EM_UNDO, 0, 0)
    Call PostMessage(Text1(1).hWnd, EM_UNDO, 0, 0)
    Call PostMessage(Text1(2).hWnd, EM_UNDO, 0, 0)
End Sub

'=======================================================
'clear le buffer des Undo
'=======================================================
Public Sub ClearUndoBuffer()
    Call SendMessage(Text1(0).hWnd, EM_EMPTYUNDOBUFFER, 0, 0)
    Call SendMessage(Text1(1).hWnd, EM_EMPTYUNDOBUFFER, 0, 0)
    Call SendMessage(Text1(2).hWnd, EM_EMPTYUNDOBUFFER, 0, 0)
End Sub

'=======================================================
'effectue un Copier de la sélection
'=======================================================
Public Sub Copy()
    Call PostMessage(Text1(0).hWnd, WM_COPY, 0, 0)
    Call PostMessage(Text1(1).hWnd, WM_COPY, 0, 0)
    Call PostMessage(Text1(2).hWnd, WM_COPY, 0, 0)
End Sub

'=======================================================
'effectue un Couper de la sélection
'=======================================================
Public Sub Cut()
    Call PostMessage(Text1(0).hWnd, WM_CUT, 0, 0)
    Call PostMessage(Text1(1).hWnd, WM_CUT, 0, 0)
    Call PostMessage(Text1(2).hWnd, WM_CUT, 0, 0)
End Sub

'=======================================================
'effectue un Coller de la sélection
'=======================================================
Public Sub Paste()
    Call PostMessage(Text1(0).hWnd, WM_PASTE, 0, 0)
    Call PostMessage(Text1(1).hWnd, WM_PASTE, 0, 0)
    Call PostMessage(Text1(2).hWnd, WM_PASTE, 0, 0)
End Sub

'=======================================================
'Scrolls the caret into view.
'=======================================================
Public Sub EnsureCaretVisibility()
    Call SendMessage(Text1(0).hWnd, EM_SCROLLCARET, 0&, 0&)
    Call SendMessage(Text1(1).hWnd, EM_SCROLLCARET, 0&, 0&)
    Call SendMessage(Text1(2).hWnd, EM_SCROLLCARET, 0&, 0&)
End Sub

'=======================================================
'récupère la ligne en x,y
'=======================================================
Public Function GetLineByCoords(ByVal x, ByVal y) As Long
Dim LineIndex  As Long
    x = Int(x / 15): y = Int(y / 15)
    LineIndex = SendMessage(Text1(lCurrentTextBox).hWnd, EM_CHARFROMPOS, 0, CLng(y * 2 ^ 16 + x))
    GetLineByCoords = LineIndex \ 2 ^ 16
End Function

'=======================================================
'récupère le caractère en x,y
'=======================================================
Public Function GetCharByCoords(ByVal x, ByVal y) As Long
Dim CharIndex  As Long
    x = Int(x / 15): y = Int(y / 15)
    CharIndex = SendMessage(Text1(lCurrentTextBox).hWnd, EM_CHARFROMPOS, 0, CLng(y * &H10000 + x))
    GetCharByCoords = Int(CharIndex And &HFFFF&)
End Function

'=======================================================
'récupère le caractère n
'=======================================================
Public Function GetChar(ByVal n As Long) As String
Dim bt As Byte

    If n <= 0 Or n > Len(sText) Then Exit Function
    
    'on fait du copymemory pour aller plus vite que le Mid
    Call CopyMemory(bt, ByVal (StrPtr(sText) + n), 1)
    
    GetChar = Chr$(bt)
End Function

'=======================================================
'récupère la ligne n
'=======================================================
Public Function GetLine(ByVal n As Long) As String
Dim Buf As String
Dim lPos As Long
Dim lLen As Long
    
    If n <= 0 Or n > Me.LineCount Then Exit Function
    
    n = n - 1   'la première ligne c'est 0
    
    'récupère le premier char de la ligne
    lPos = Me.FirstChar(n)
    
    'récupère la taille de la ligne en donnant son premier char
    lLen = SendMessage(Text1(lCurrentTextBox).hWnd, EM_LINELENGTH, lPos, 0&)
    
    'dimensionne le buffer
    Buf = Space$(lLen)
    
    'récupère le texte
    Call SendMessageStr(Text1(lCurrentTextBox).hWnd, EM_GETLINE, n, Buf)
    
    GetLine = Buf

End Function

'=======================================================
'ajoute une line à la suite
'=======================================================
Public Sub AppendLine(ByVal NewLine As String)
    
    sText = sText & vbNewLine & NewLine
    Text1(lCurrentTextBox).Text = sText
    
End Sub

'=======================================================
'MAJ du controle
'=======================================================
Public Sub Refresh()
    Call RefreshPrivate
    'Call ChangeNum
End Sub
Private Sub RefreshPrivate()
Dim R As RECT
Dim st As Long
Dim hRgn As Long
Dim hBrush As Long
Dim hFont As Long
Dim ttFont As LOGFONT
Dim hObj As Long
Dim x As Long

    If bUnRefreshControl Then Exit Sub

    '//on efface
    Set UserControl.MaskPicture = Nothing
    Set UserControl.Picture = Nothing
    Call UserControl.Cls

    '//convertit les couleurs
    Call OleTranslateColor(lLegendBackColor1, 0, lLegendBackColor1)
    Call OleTranslateColor(lLegendBackColor2, 0, lLegendBackColor2)
    Call OleTranslateColor(lForeColor, 0, lForeColor)
    Call OleTranslateColor(lBorderColor, 0, lBorderColor)
    Call OleTranslateColor(lLegendForeColor, 0, lLegendForeColor)
    

    '//affiche le BackGround de la zone Legend
    If tLegendType <> NoLegend Then
        If tLegendGradient = None Then
            'pas de gradient
            'on dessine alors un rectangle
            Line (15, 15)-(lLegendWidth, ScaleHeight - 30), lLegendBackColor1, BF
        ElseIf tLegendGradient = Horizontal Then
            'gradient horizontal
            Call FillGradient(UserControl.hDC, lLegendBackColor1, lLegendBackColor2, _
                lLegendWidth / Screen.TwipsPerPixelX, Height / _
                Screen.TwipsPerPixelY, Horizontal)
        Else
            'gradient vertical
            Call FillGradient(UserControl.hDC, lLegendBackColor1, lLegendBackColor2, _
                lLegendWidth / Screen.TwipsPerPixelX, Height / _
                Screen.TwipsPerPixelY, Vertical)
        End If
        Line (lLegendWidth, 15)-(lLegendWidth, ScaleHeight - 15), lBorderColor
    End If


    '//on dessine donc maintenant le bord
    If bDisplayBorder Then
        'alors c'est un rectangle
        
        'on défini un brush
        hBrush = CreateSolidBrush(lBorderColor)
        
        'on définit une zone rectangulaire à bords arrondi
        hRgn = CreateRectRgn(0, 0, ScaleWidth / Screen.TwipsPerPixelX, _
            ScaleHeight / Screen.TwipsPerPixelY)
        
        'on dessine le contour
        Call FrameRgn(UserControl.hDC, hRgn, hBrush, 1, 1)
    
        '//on sauve la picture dans le MaskPicture
        'sera utilisée lors du réaffichage des lignes pour éviter
        'de tout retracer à chaque fois
        Set UserControl.MaskPicture = UserControl.Image
        
        'on détruit le brush et la zone
        Call DeleteObject(hBrush)
        Call DeleteObject(hRgn)
    End If


    '//maintenant on dessine le texte de la zone Legend
    If tLegendType = SimpleTextLegend Then
        'texte simple
        
        If tLegendOrientation = VerticalLegend Then
            'alors vertical
            
            'récupère la police utilisée (celle du Usercontrol)
            hFont = SelectObject(hDC, CreateFontIndirect(ttFont))
            
            'récupère un LOGFONT
            Call GetObjectA(hFont, Len(ttFont), ttFont)
            
            'change l'angle
            ttFont.lfEscapement = 900   '10 * 90°
            
            'on associe la font au DC du UserControl
            hObj = CreateFontIndirect(ttFont)
            Call SelectObject(hDC, hObj)
    
            UserControl.ForeColor = lLegendForeColor
            
            'on centre la police horizontalement
            If tLegendAlignmentX = vbCenter Then
                CurrentX = Int(lLegendWidth - TextHeight(sLegendText)) / 2
            ElseIf tLegendAlignmentX = vbLeftJustify Then
                CurrentX = 0
            Else
                CurrentX = lLegendWidth - 30 - TextHeight(sLegendText)
            End If
            
            'on définit la hauteur en fonction de l'alignement
            If tLegendAlignmentY = vbCenter Then
                CurrentY = (Height - TextWidth(sLegendText)) / 2 '+ TextWidth(sLegendText)
            ElseIf tLegendAlignmentY = vbLeftJustify Then
                CurrentY = 0
            Else
                CurrentY = Height - TextWidth(sLegendText) - 45
            End If
            
            'on affiche le texte
            Call SetRect(R, CurrentX / Screen.TwipsPerPixelX, _
                ScaleHeight / 15 - 2 - CurrentY / Screen.TwipsPerPixelY, _
                lLegendWidth / Screen.TwipsPerPixelX, -ScaleHeight / 15)
            
            If bUnicode Then
                Call DrawTextW(UserControl.hDC, StrPtr(sLegendText), Len(sLegendText), R, DT_LEFT)
            Else
                Call DrawText(UserControl.hDC, sLegendText, Len(sLegendText), R, DT_LEFT)
            End If
            
            'on vire les objets créés
            Call DeleteObject(hObj)
        Else
            'alors horizontal
            
            'récupère la police utilisée (celle du Usercontrol)
            hFont = SelectObject(hDC, CreateFontIndirect(ttFont))
            
            'récupère un LOGFONT
            Call GetObjectA(hFont, Len(ttFont), ttFont)
            
            'change l'angle
            ttFont.lfEscapement = 0
            
            'on associe la font au DC du UserControl
            hObj = CreateFontIndirect(ttFont)
            Call SelectObject(hDC, hObj)
    
            UserControl.ForeColor = lLegendForeColor
            
            'on centre la police horizontalement
            If tLegendAlignmentX = vbCenter Then
                CurrentX = Int(lLegendWidth - TextWidth(sLegendText)) / 2
            ElseIf tLegendAlignmentX = vbLeftJustify Then
                CurrentX = 30
            Else
                CurrentX = lLegendWidth - 15 - TextWidth(sLegendText)
            End If
            
            'on définit la hauteur en fonction de l'alignement
            If tLegendAlignmentY = vbCenter Then
                CurrentY = (Height - TextHeight(sLegendText)) / 2
            ElseIf tLegendAlignmentY = vbLeftJustify Then
                CurrentY = 0
            Else
                CurrentY = ScaleHeight - TextHeight(sLegendText) - 30
            End If
            
            'on affiche le texte
            Call SetRect(R, CurrentX / Screen.TwipsPerPixelX, CurrentY / Screen.TwipsPerPixelY, lLegendWidth / Screen.TwipsPerPixelX, Height / Screen.TwipsPerPixelY)
            
            If bUnicode Then
                Call DrawTextW(UserControl.hDC, StrPtr(sLegendText), Len(sLegendText), R, DT_LEFT)
            Else
                Call DrawText(UserControl.hDC, sLegendText, Len(sLegendText), R, DT_LEFT)
            End If
            
            'on vire les objets créés
            Call DeleteObject(hObj)
            
        End If
        
    ElseIf tLegendType = LineNumberLegend Then
        'on affiche les lignes
        
        'Call ChangeNum
        
        
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

Private Sub VS_Change(Value As Currency)
Dim zNumber As Long
'Static lngOldValue As Currency

    bDamn = True
    
    'récupère la première ligne visible
    zNumber = FirstVisibleLine
    
    'on change maintenant la FirstVisibleLine
    Call SendMessage(Text1(lCurrentTextBox).hWnd, EM_LINESCROLL, 0, _
        CLng(Value - zNumber - 1))
        
    'on en refresh QUE si on a changé de value entre temps
   ' If lngOldValue <> CLng(Value) Then Call ChangeNum
   ' lngOldValue = Value

    'bDamn = False
    RaiseEvent VScrolled(Value)
End Sub

Private Sub VS_Scroll()
    Call VS_Change(VS.Value)
End Sub

'=======================================================
'change la numérotation
'=======================================================
Public Sub RefreshNum(Optional ByVal Force As Boolean = False)
    Call ChangeNum(Force)
End Sub
Private Sub ChangeNum(Optional ByVal Force As Boolean = False)
Dim x As Long
Dim z As Long
Dim st As Long
Static FirstV As Long
Static d As Long
Dim F As Long
    
    If bUnRefreshControl Then Exit Sub
    
    'si pas de 'Force', alors on vérifie que l'on doit bien changer de Num
    If Force = False Then
        If bNot3 = 1 Then
            bNot3 = 2
            Exit Sub
        End If
        
        'on ne change les Nums QUE si le FirstVisible a changé
        F = FirstVisibleLine
        If FirstV = F Then Exit Sub
        FirstV = F
    End If
   ' d = d + 1
   ' Debug.Print d

    If tLegendType = LineNumberLegend Then
        'alors il faut refresh les numéros de ligne
        
        'on récupère la hauteur d'une ligne DE LA TEXTBOX
        st = lCharH * 15
        
        z = Int(Text1(1).Height / st)
    
        '//on clear le controle
        Call UserControl.Cls
        
        '//on pose la maskpicture en picture
        Set UserControl.Picture = UserControl.MaskPicture
        
        'ajuste la taille si nécessaire
        If bLegendAutoSize Then
            x = TextWidth(CStr(LineCount - 1)) + 60
            If x > lLegendWidth Then
                Set UserControl.MaskPicture = Nothing
                lLegendWidth = x
                Call UserControl_Resize
            End If
        End If
        
        
        '//on retrace le texte
        
        UserControl.ForeColor = lLegendForeColor
      
        For x = FirstVisibleLine To FirstVisibleLine + z - 1
            If x < LineCount Then

                'on change le CurrentX selon l'alignment désiré
                If tLegendAlignmentX = vbCenter Then
                    CurrentX = Int(lLegendWidth - TextWidth(CStr(x + 1))) / 2
                ElseIf tLegendAlignmentX = vbLeftJustify Then
                    CurrentX = 30
                Else
                    CurrentX = lLegendWidth - 15 - TextWidth(CStr(x + 1))
                End If

                'on change le CurrentY
                CurrentY = (x - FirstVisibleLine) * st + 60

                'on Print
                Print CStr(x + 1)
            End If
        Next x
    End If
End Sub

Private Sub HS_Change(Value As Currency)
Dim zNumber As Long
Dim vale As Long
Static lngOldValue As Currency
    
    If bOk4 Then Exit Sub

   ' bDoNotSync = True
    
    'on change maintenant la FirstVisibleLine
    vale = GetScrollPos(Text1(2).hWnd, SB_HORZ) 'value de la HScroll classique
    
    If Abs(Value - lngOldValue) < 1 Then Exit Sub
    
    'scroll horizontalement
    Call SendMessage(Text1(2).hWnd, EM_LINESCROLL, Value - lngOldValue, 0)
    If Value = 0 Then
        'alors tout à gauche
        Call SendMessage(Text1(2).hWnd, EM_LINESCROLL, -1, 0)
    ElseIf Round(Value) = HS.Max Then
        'alors tout à droite
        Call SendMessage(Text1(2).hWnd, EM_LINESCROLL, 1, 0)
    End If
    
    lngOldValue = Value

    RaiseEvent HScrolled(Value)
    
    Call SyncHS
End Sub

Private Sub HS_MouseUp(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    bDoNotSync = False
End Sub

Private Sub HS_Scroll()
    bOk4 = False: bDoNotSync = True
    Call HS_Change(HS.Value)
    bOk4 = True
End Sub

'=======================================================
'récupère le Max de la HScroll classique
'=======================================================
Private Function GetHMax()
Dim lPsi As SCROLLINFO
    lPsi.cbSize = Len(lPsi)
    lPsi.fMask = SIF_RANGE
    Call GetScrollInfo(Text1(2).hWnd, SB_HORZ, lPsi)
    GetHMax = lPsi.nMax
End Function

'=======================================================
'récupère le Value de la HScroll classique
'=======================================================
Private Function GetHValue()
Dim lPsi As SCROLLINFO
    lPsi.cbSize = Len(lPsi)
    lPsi.fMask = SIF_ALL
    Call GetScrollInfo(Text1(2).hWnd, SB_HORZ, lPsi)
    GetHValue = lPsi.nTrackPos
End Function

'=======================================================
'synchronise les HSCrolls perso et classiques
'=======================================================
Private Sub SyncHS()
Dim lPsi As SCROLLINFO
    
    If (tScrollBars And vbVertical) <> vbVertical Or bDoNotSync Then Exit Sub
    
    lPsi.cbSize = Len(lPsi)
    lPsi.fMask = SIF_ALL
    Call GetScrollInfo(Text1(2).hWnd, SB_HORZ, lPsi)
    
    With HS
        .Max = Int((lPsi.nMax - lPsi.nPage) / 7)
        bOk4 = True
        .Value = lPsi.nTrackPos / 7
        
        bOk4 = False
    End With
   ' Debug.Print "sinc  " & HS.Value
End Sub

Public Function gety() As String
    gety = GetHValue & "  /  " & GetHMax
End Function


