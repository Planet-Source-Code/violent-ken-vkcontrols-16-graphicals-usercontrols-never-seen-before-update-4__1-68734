VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
'ENUMS
'=======================================================
Public Enum BackStyleConstants
    TRANSPARENT = 0
    Opaque = 1
End Enum
Public Enum Label_Mode
    No = 0
    PercentageMode = 1
    ValueMode = 2
    Steps = 3
End Enum
Public Enum Border
    NoBorder = 0
    FixedSingle = 1
End Enum
Public Enum Mode_Degrade
    OnlyValue = 1
    AllLengh = 2
End Enum
Public Enum Button_Type
    NoButton = 0
    LeftButton = 1
    RightButton = 2
    MiddleButton = 4
End Enum
Public Enum Text_Alignment
    TopLeft = 1
    TopCenter = 2
    TopRight = 3
    MiddleLeft = 4
    MiddleCenter = 5
    MiddleRight = 6
    BottomLeft = 7
    BottomCenter = 8
    BottomRight = 9
End Enum
Public Enum Wheel_Sens
    WHEEL_UP
    WHEEL_DOWN
End Enum
Public Enum GradientConstants
    None = 0
    Vertical = 1
    Horizontal = 2
End Enum
Public Enum PictureAlignment
    [Left Justify]
    [Right Justify]
End Enum
Public Enum Direction
    [Left_Right]
    [Up_Down]
End Enum
Public Enum SortOrder
    [Alphabetical]
    [InverseAlphabetical]
    [DoNotSort]
End Enum
Public Enum ListBoxType
    [SimpleList]
    [FileListBox]
    [FolderListBox]
    [DriveListBox]
End Enum
Public Enum FileAttributes
    Alias = 1024
    Archive = 32
    Compressed = 2048
    Directory = 16
    Hidden = 2
    Normal = 0
    ReadOnly = 1
    System = 4
    Volume = 8
    All = Alias Or Archive Or Compressed Or Hidden Or ReadOnly Or System
End Enum
Public Enum LegendTypeConstants
    [NoLegend]
    [SimpleTextLegend]
    [LineNumberLegend]
End Enum
Public Enum OrientationsConstants
    [HorizontalLegend] = 0
    [VerticalLegend] = 1
End Enum
Public Enum vkCommandStyle
    [No Style - vkCommand] = 0
    [Defaut Style - vkCommand] = 1
    [XP Style - vkCommand] = 2
    [Office Style - vkCommand] = 3
    [Plastik Style - vkCommand] = 4
    [Galaxy Style - vkCommand] = 5
    [JAVA Style - vkCommand] = 6
End Enum
