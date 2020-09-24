VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vkListItem"
Attribute VB_GlobalNameSpace = True
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

Public tagString1 As String
Attribute tagString1.VB_VarDescription = "String variable to store informations about item"
Public tagString2 As String
Attribute tagString2.VB_VarDescription = "String variable to store informations about item"
Public tagLong As Long
Attribute tagLong.VB_VarDescription = "Long variable to store informations about item"


Private lBackColor As Long
Private bChecked As Boolean
Private tFont As StdFont
Private lForeColor As Long
Private tIcon As Variant
Private lIndex As Long
Private sKey As String
Private bSelected As Boolean
Private sText As String
Private lHeight As Long
Private tAlig As AlignmentConstants
Private lpxlIconHeight As Long
Private lpxlIconWidth As Long
Private lSelColor As Long
Private lBorderSelColor As Long
Private lpctType As Long


'=======================================================
'properties
'=======================================================
Public Property Get BackColor() As Long: BackColor = lBackColor: End Property
Attribute BackColor.VB_Description = "Back color"
Public Property Let BackColor(BackColor As Long): lBackColor = BackColor: End Property
Public Property Get Checked() As Boolean: Checked = bChecked: End Property
Attribute Checked.VB_Description = "Get if item is checked or not"
Public Property Let Checked(Checked As Boolean): bChecked = Checked: End Property
Public Property Get Font() As StdFont: Set Font = tFont: End Property
Attribute Font.VB_Description = "Item font"
Public Property Let Font(Font As StdFont): Set tFont = Font: End Property
Public Property Get ForeColor() As Long: ForeColor = lForeColor: End Property
Attribute ForeColor.VB_Description = "Text color"
Public Property Let ForeColor(ForeColor As Long): lForeColor = ForeColor: End Property
Public Property Get Icon() As Variant:  Icon = tIcon: End Property
Attribute Icon.VB_Description = "Icon to display"
Public Property Let Icon(Icon As Variant):  tIcon = Icon: End Property
Public Property Get Index() As Long: Index = lIndex: End Property
Attribute Index.VB_Description = "Index of the item"
Public Property Let Index(Index As Long): lIndex = Index: End Property
Public Property Get Key() As String: Key = sKey: End Property
Attribute Key.VB_Description = "Key"
Public Property Let Key(Key As String): sKey = Key: End Property
Public Property Get Selected() As Boolean: Selected = bSelected: End Property
Attribute Selected.VB_Description = "Get if item is selected"
Public Property Let Selected(Selected As Boolean): bSelected = Selected: End Property
Public Property Get Text() As String: Text = sText: End Property
Attribute Text.VB_Description = "Text to display"
Public Property Let Text(Text As String): sText = Text: End Property
Public Property Get Height() As Long: Height = lHeight: End Property
Attribute Height.VB_Description = "Height of the item"
Public Property Let Height(Height As Long): lHeight = Height: End Property
Public Property Get Alignment() As AlignmentConstants: Alignment = tAlig: End Property
Attribute Alignment.VB_Description = "Text alignment"
Public Property Let Alignment(Alignment As AlignmentConstants): tAlig = Alignment: End Property
Public Property Get pxlIconHeight() As Long: pxlIconHeight = lpxlIconHeight: End Property
Attribute pxlIconHeight.VB_Description = "Height of icon (pixels)"
Public Property Let pxlIconHeight(pxlIconHeight As Long): lpxlIconHeight = pxlIconHeight: End Property
Public Property Get pxlIconWidth() As Long: pxlIconWidth = lpxlIconWidth: End Property
Attribute pxlIconWidth.VB_Description = "Width of icon (pixels)"
Public Property Let pxlIconWidth(pxlIconWidth As Long): lpxlIconWidth = pxlIconWidth: End Property
Public Property Get SelColor() As Long: SelColor = lSelColor: End Property
Attribute SelColor.VB_Description = "Selection color"
Public Property Let SelColor(SelColor As Long): lSelColor = SelColor: End Property
Public Property Get BorderSelColor() As Long: BorderSelColor = lBorderSelColor: End Property
Attribute BorderSelColor.VB_Description = "Border color of the item when selected"
Public Property Let BorderSelColor(BorderSelColor As Long): lBorderSelColor = BorderSelColor: End Property
Public Property Get pctType() As Long: pctType = lpctType: End Property
Public Property Let pctType(pctType As Long): lpctType = pctType: End Property


Private Sub Class_Initialize()
'on affecte les valeurs par défaut
    With Me
        .BackColor = &HFFFFFF
        .ForeColor = vbBlack
        .Alignment = vbLeftJustify
        .SelColor = 16768444
        '.Font =
        .BorderSelColor = 16419097
        .pxlIconHeight = 16
        .pxlIconWidth = 16
        .Height = 255
        .pctType = 0
    End With
End Sub
