VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vkPrivateScroll"
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


Private bCol As OLE_COLOR
Private lArrowColor As OLE_COLOR
Private lFrontColor As OLE_COLOR
Private lBorderColor As OLE_COLOR
Private lSelColor As OLE_COLOR
Private lLargeChangeColor As OLE_COLOR
Private lDownColor As OLE_COLOR
Private bEnable As Boolean
Private bNotOk As Boolean
Private bNotOk2 As Boolean
Private lScrollHeight As Byte
Private lMin As Currency
Private lMax As Currency
Private lValue As Currency
Private lSmallChange As Currency
Private lLargeChange As Currency
Private lWheelChange As Currency
Private bEnableWheel As Boolean
Private lGrise As Long
Private lMouseInterval As Long
Private bUnRefreshControl As Boolean
Private lhDc As Long
Private lhWnd As Long
Private lWidth As Long
Private lHeight As Long

Public Property Get hDc() As Long: hDc = lhDc: End Property
Public Property Get hwnd() As Long: hwnd = lhWnd: End Property
Friend Property Let hDc(hDc As Long): lhDc = hDc: End Property
Friend Property Let hwnd(hwnd As Long): lhWnd = hwnd: End Property
Public Property Get BackColor() As OLE_COLOR: BackColor = bCol: End Property
Public Property Let BackColor(BackColor As OLE_COLOR): bCol = BackColor: End Property
Public Property Get BorderColor() As OLE_COLOR: BorderColor = lBorderColor: End Property
Public Property Let BorderColor(BorderColor As OLE_COLOR): lBorderColor = BorderColor: End Property
Public Property Get ArrowColor() As OLE_COLOR: ArrowColor = lArrowColor: End Property
Public Property Let ArrowColor(ArrowColor As OLE_COLOR): lArrowColor = ArrowColor: End Property
Public Property Get FrontColor() As OLE_COLOR: FrontColor = lFrontColor: End Property
Public Property Let FrontColor(FrontColor As OLE_COLOR): lFrontColor = FrontColor: End Property
Public Property Get Enabled() As Boolean: Enabled = bEnable: End Property
Public Property Let Enabled(Enabled As Boolean): bEnable = Enabled: End Property
Public Property Get EnableWheel() As Boolean: EnableWheel = bEnableWheel: End Property
Public Property Let EnableWheel(EnableWheel As Boolean): bEnableWheel = EnableWheel: End Property
Public Property Get ScrollHeight() As Byte: ScrollHeight = lScrollHeight: End Property
Public Property Let ScrollHeight(ScrollHeight As Byte): lScrollHeight = ScrollHeight: End Property
Public Property Get Min() As Currency: Min = lMin: End Property
Friend Property Let Min(Min As Currency): lMin = Min: End Property
Public Property Get Max() As Currency: Max = lMax: End Property
Friend Property Let Max(Max As Currency): lMax = Max: End Property
Public Property Get Value() As Currency: Value = lValue: End Property
Public Property Let Value(Value As Currency): lValue = Value: End Property
Public Property Get SmallChange() As Currency: SmallChange = lSmallChange: End Property
Public Property Let SmallChange(SmallChange As Currency): lSmallChange = SmallChange: End Property
Public Property Get LargeChange() As Currency: LargeChange = lLargeChange: End Property
Public Property Let LargeChange(LargeChange As Currency): lLargeChange = LargeChange: End Property
Public Property Get WheelChange() As Currency: WheelChange = lWheelChange: End Property
Public Property Let WheelChange(WheelChange As Currency): lWheelChange = WheelChange: End Property
Public Property Get DownColor() As OLE_COLOR: DownColor = lDownColor: End Property
Public Property Let DownColor(DownColor As OLE_COLOR): lDownColor = DownColor: End Property
Public Property Get MouseHoverColor() As OLE_COLOR: MouseHoverColor = lSelColor: End Property
Public Property Let MouseHoverColor(MouseHoverColor As OLE_COLOR): lSelColor = MouseHoverColor: End Property
Public Property Get MouseInterval() As Long: MouseInterval = lMouseInterval: End Property
Public Property Let MouseInterval(MouseInterval As Long): lMouseInterval = MouseInterval: End Property
Public Property Get LargeChangeColor() As OLE_COLOR: LargeChangeColor = lLargeChangeColor: End Property
Public Property Let LargeChangeColor(LargeChangeColor As OLE_COLOR): lLargeChangeColor = LargeChangeColor: End Property
Public Property Get UnRefreshControl() As Boolean: UnRefreshControl = bUnRefreshControl: End Property
Public Property Let UnRefreshControl(UnRefreshControl As Boolean): bUnRefreshControl = UnRefreshControl: End Property
Public Property Get Width() As Long: Width = lWidth: End Property
Public Property Let Width(Width As Long): lWidth = Width: End Property
Friend Property Let Height(Height As Long): lHeight = Height: End Property
Public Property Get Height() As Long: Height = lHeight: End Property

Private Sub Class_Initialize()
    With Me
        .ArrowColor = vbWhite
        .BackColor = vbWhite
        .BorderColor = &HFF8080
        .Enabled = True
        .EnableWheel = True
        .FrontColor = 15782079
        .LargeChange = 5
        .ScrollHeight = 25
        .SmallChange = 1
        .WheelChange = 3
        .DownColor = 12492429
        .MouseHoverColor = vbWhite
        .MouseInterval = 50
        .LargeChangeColor = 12492429
        .UnRefreshControl = True
    End With
End Sub
