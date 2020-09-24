Attribute VB_Name = "mdlGray"
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
'CONSTANTES
'=======================================================
Private Const BI_RGB                        As Long = 0&
Private Const DIB_RGB_COLORS                As Long = 0


'=======================================================
'TYPES
'=======================================================
Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type


'=======================================================
'APIS
'=======================================================
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetDIBColorTable Lib "gdi32" (ByVal hDc As Long, ByVal un1 As Long, ByVal un2 As Long, pRGBQuad As RGBQUAD) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hDc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long


'=======================================================
'convertit une image contenu dans une PictureBox en niveaux de gris
'=======================================================
Public Sub GrayScale(PicSRC As PictureBox)

Const pixR = 1
Const pixG = 2
Const pixB = 3

Dim bitmap_info As BITMAPINFO
Dim pixels() As Byte
Dim bytes_per_scanLine As Integer
Dim x As Integer
Dim y As Integer
Dim ave_color As Byte
Dim bw As Long
Dim bh As Long

    bw = PicSRC.ScaleWidth
    bh = PicSRC.ScaleHeight
    
    bytes_per_scanLine = ((((bw * 32) + 31) \ 32) * 4)
  
  ' Prepare la bitmap description.
    With bitmap_info.bmiHeader
        .biSize = 40
        .biWidth = bw
        .biHeight = -bh
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
        .biSizeImage = bytes_per_scanLine * bh
         End With
  ' Transforme en bitmap's data.
    ReDim pixels(1 To 4, 1 To bw, 1 To bh)
    Call GetDIBits(PicSRC.hDc, PicSRC.Image, 0, bh, pixels(1, 1, 1), _
        bitmap_info, DIB_RGB_COLORS)
  ' Modifie les pixels.
    For y = 1 To bh
        For x = 1 To bw
            ave_color = CByte((CInt(pixels(pixR, x, y)) + pixels(pixG, x, y) + _
                pixels(pixB, x, y)) \ 3)
         '  une autre possibilité:
         '  ave_color = CByte((CInt(pixels(pixR, x, y)) * 0.299 + pixels(pixG, x, y) * 0.587 + pixels(pixB, x, y) * 0.114))
            pixels(pixR, x, y) = ave_color
            pixels(pixG, x, y) = ave_color
            pixels(pixB, x, y) = ave_color
            Next x
        Next y
  ' Affiche le resultat.
    Call SetDIBits(PicSRC.hDc, PicSRC.Image, 0, bh, pixels(1, 1, 1), _
        bitmap_info, DIB_RGB_COLORS)
        
    PicSRC.Picture = PicSRC.Image
    
End Sub
