VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
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


'=======================================================
'CONSTANTES
'=======================================================
Private Const FILE_ATTRIBUTE_DIRECTORY      As Long = &H10
Private Const PICTYPE_ICON                  As Long = 3
Private Const SHGFI_LARGEICON               As Long = &H0
Private Const SHGFI_SMALLICON               As Long = &H1
Private Const SHGFI_ICON                    As Long = &H100


'=======================================================
'TYPES & ENUMS
'=======================================================
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type
Public Enum IconSize
    Size16 = 16
    Size32 = 32
End Enum
Private Type TypeIcon
    cbSize As Long
    picType As PictureTypeConstants
    hIcon As Long
End Type
Private Type CLSID
    ID((123)) As Byte
End Type


'=======================================================
'APIS
'=======================================================
Private Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function SHGetFileInfo Lib "Shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal Flags&) As Long
Private Declare Function ExtractAssociatedIcon Lib "Shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Boolean
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As TypeIcon, riid As CLSID, ByVal fown As Long, lpUnk As Object) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function PathFindExtension Lib "shlwapi" Alias "PathFindExtensionA" (ByVal pPath As String) As Long
Private Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal ptr As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal ptr As Any) As Long


'=======================================================
'PUBLIC FUNCTIONS
'=======================================================

'=======================================================
'Enumère les fichiers du dossier
'=======================================================
Public Sub EnumFiles(ByVal Folder As String, ByRef Result() As String, _
    Optional ByVal SubFolder As Boolean = False, Optional ByVal bSystem As _
    Boolean = True, Optional ByVal bHidden As Boolean = True, Optional _
    ByVal bReadOnly As Boolean = True, Optional ByVal Pattern As String _
    = vbNullString)
    
Dim x As Long
'Dim s() As String

    'ajoute un slash si nécessaire
    If Right$(Folder, 1) <> "\" Then Folder = Folder & "\"

    'lance l'énumération
    ReDim Result(0)
    Call EnumFilesFromFolder(Folder, Result(), True, SubFolder, bSystem, _
        bReadOnly, bHidden, Pattern)
    
    'donne le résultat
   ' Result = s
    
End Sub

'=======================================================
'Enumère les sous dossiers du dossier
'=======================================================
Public Sub EnumFolders(ByVal Folder As String, ByRef Result() As _
    String, Optional ByVal SubFolder As Boolean = False, _
    Optional ByVal bSystem As Boolean = True, Optional ByVal bHidden _
    As Boolean = True, Optional ByVal bReadOnly As Boolean = True)
    
Dim x As Long
'Dim s() As String

    If Right$(Folder, 1) <> "\" Then Folder = Folder & "\"
    
    'lance l'énumération
    ReDim Result(0)
    Call EnumFoldersPrivate(Folder, Result(), True, SubFolder, bSystem, bReadOnly, bHidden)
    
    'résultat
    'Result = s
    
End Sub

'=======================================================
'Enumère tous les drives présents
'=======================================================
Public Sub EnumDrives(ByRef Result() As String)
Dim strBuffer As String
Dim x As Byte, lngLen As Long
Dim s() As String

    'créé un buffer qui recevra les noms de partition
    strBuffer = Space$(255)
    
    'obtient les noms
    lngLen = GetLogicalDriveStrings(255, strBuffer)
    
    'formate la string
    strBuffer = Left$(strBuffer, lngLen)
    
    'redimensionne le tableau
    lngLen = (lngLen \ 4) - 1
    
    'redimensionne le tableau
    ReDim s(lngLen + 1)
    
    For x = 1 To lngLen + 1
        s(x) = Mid$(strBuffer, 4 * x - 3, 3)
    Next x
    
    Result = s
End Sub

'=======================================================
'Récpère une icone depuis un fichier/dossier/drive
'=======================================================
Public Function GetIcon(ByVal Source As String, ByVal Size As IconSize) _
    As StdPicture
    
Dim SHFinfo As SHFILEINFO
Dim lFlag As Long
Dim tCLSID As CLSID
Dim tTypeIcon As TypeIcon
Dim hIcon As Long
 
    'type de flag en fonction de la taille
    If Size = Size16 Then lFlag = SHGFI_SMALLICON Else _
        lFlag = SHGFI_LARGEICON
    
    'on récupère les infos sur le fichier
    hIcon = SHGetFileInfo(Source, 0&, SHFinfo, Len(SHFinfo), SHGFI_ICON + lFlag)
    
    'initialise les structures pour récupérer l'image
    With tTypeIcon
        .cbSize = Len(tTypeIcon)
        .picType = vbPicTypeIcon
        .hIcon = SHFinfo.hIcon
    End With
    With tCLSID
        .ID(8) = &HC0
        .ID(15) = &H46
    End With
    
    'créé la picture
    Call OleCreatePictureIndirect(tTypeIcon, tCLSID, 1, GetIcon)
    
    'détruit l'icone pour libérer la mémoire
    Call DestroyIcon(hIcon)
End Function



'=======================================================
'PRIVATE SUB AND FUNCTIONS
'=======================================================

'=======================================================
'Enumère les fichiers d'un dossier
'De 1 à ubound
'=======================================================
Private Sub EnumFilesFromFolder(ByVal Directory As String, Files() As String, _
    Optional Begin As Boolean = False, Optional SubFolder As Boolean _
    = True, Optional ByVal bSystem As Boolean = True, Optional ByVal _
    bReadOnly As Boolean = True, Optional ByVal bHidden As Boolean = True, _
    Optional ByVal Pattern As String = vbNullString)
    
'infos d'un fichier et handle d'énumération
Dim FileInfo As WIN32_FIND_DATA, hFind As Long
'limite de Files
Static ub As Long

    'si on commence la limite est 0, pas d'enregistrement dans le tableau
    If Begin = True Then ub = 0
    
    'ouvre le dossier pour récupérer la liste de ses fichiers
    hFind = FindFirstFile(Directory & "*", FileInfo)
    If hFind <> -1 Then
        'si le fichier est un dossier
        If (FileInfo.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            'si ce n'est ni . ni ..
            If InStr(FileInfo.cFileName, ".") <> 1 And SubFolder Then
                'on récupère la liste des fichiers de ce sous dossier
                EnumFilesFromFolder Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1) & "\", Files, False, , bSystem, bReadOnly, bHidden
            End If
        'sinon on l'ajoute à la fin de la liste
        Else
            'vérifie l'attribut
            If Not (bHidden) And (FileInfo.dwFileAttributes And Hidden) = Hidden Or _
                Not (bSystem) And (FileInfo.dwFileAttributes And System) = System Or _
                Not (bReadOnly) And (FileInfo.dwFileAttributes And ReadOnly) = ReadOnly Or CheckPattern(Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1), Pattern) = False Then
                
            Else
                ub = ub + 1
                ReDim Preserve Files(ub)
                'chemin complet
                Files(ub) = Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1)
            End If
        End If
        'tant qu'il y a des fichiers dans le dossier en cours
        Do While FindNextFile(hFind, FileInfo)
            'on peut arrêter à tout moment
            DoEvents
            
            'si le fichier est un dossier
            If (FileInfo.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
                'si ce n'est ni . ni ..
                If InStr(FileInfo.cFileName, ".") <> 1 And SubFolder Then
                    'on récupère la liste des fichiers de ce sous dossier
                    EnumFilesFromFolder Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1) & "\", Files, False, , bSystem, bReadOnly, bHidden
                End If
            'sinon on l'ajoute à la fin de la liste
            Else
                'vérifie l'attribut
                If Not (bHidden) And (FileInfo.dwFileAttributes And Hidden) = Hidden Or _
                    Not (bSystem) And (FileInfo.dwFileAttributes And System) = System Or _
                    Not (bReadOnly) And (FileInfo.dwFileAttributes And ReadOnly) = ReadOnly Or CheckPattern(Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1), Pattern) = False Then
                    
                Else
                    ub = ub + 1
                    ReDim Preserve Files(ub)
                    'chemin complet
                    Files(ub) = Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1)
                End If
            End If
        Loop
    End If
    'ferme l'énumèration
    FindClose hFind
End Sub

'=======================================================
'Enumère les sous dossiers d'un dossier
'De 1 à ubound
'=======================================================
Private Sub EnumFoldersPrivate(ByVal Directory As String, s() As String, _
    Optional Begin As Boolean = False, Optional SubFolder As Boolean = _
    True, Optional ByVal bSystem As Boolean = True, Optional ByVal _
    bReadOnly As Boolean = True, Optional ByVal bHidden As Boolean = True)
    
'infos d'un fichier et handle d'énumération
Dim FileInfo As WIN32_FIND_DATA, hFind As Long, sDir As String
'limite de Files
Static ub As Long

    'si on commence la limite est 0, pas d'enregistrement dans le tableau
    If Begin = True Then ub = 0
    
    'ouvre le dossier pour récupérer la liste de ses fichiers
    hFind = FindFirstFile(Directory & "*", FileInfo)
    If hFind <> -1 Then
        'si le fichier est un dossier
        If (FileInfo.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            'si ce n'est ni . ni ..
            If InStr(FileInfo.cFileName, ".") <> 1 And SubFolder Then
                'on récupère la liste des fichiers de ce sous dossier
                EnumFoldersPrivate Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1) & "\", s, False, , bSystem, bReadOnly, bHidden
            End If
            
            sDir = Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1)
            If Right$(sDir, 3) <> "\.." Then
                
                'vérifie l'attribut
                If Not (bHidden) And (FileInfo.dwFileAttributes And Hidden) = Hidden Or _
                    Not (bSystem) And (FileInfo.dwFileAttributes And System) = System Or _
                    Not (bReadOnly) And (FileInfo.dwFileAttributes And ReadOnly) = ReadOnly Then
                    
                Else
                    If Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1) <> "." Then
                        ub = ub + 1
                        ReDim Preserve s(ub)
                        'chemin complet
                        s(ub) = sDir
                    End If
                End If
                
            End If
        End If
        'tant qu'il y a des fichiers dans le dossier en cours
        Do While FindNextFile(hFind, FileInfo)
            'on peut arrêter à tout moment
            DoEvents
            
            'si le fichier est un dossier
            If (FileInfo.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
                'si ce n'est ni . ni ..
                If InStr(FileInfo.cFileName, ".") <> 1 And SubFolder Then
                    'on récupère la liste des fichiers de ce sous dossier
                    EnumFoldersPrivate Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1) & "\", s, False, , bSystem, bReadOnly, bHidden
                End If
                sDir = Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1)
                If Right$(sDir, 3) <> "\.." Then
                    'vérifie l'attribut
                    If Not (bHidden) And (FileInfo.dwFileAttributes And Hidden) = Hidden Or _
                        Not (bSystem) And (FileInfo.dwFileAttributes And System) = System Or _
                        Not (bReadOnly) And (FileInfo.dwFileAttributes And ReadOnly) = ReadOnly Then
                        
                    Else
                        If Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1) <> "." Then
                            ub = ub + 1
                            ReDim Preserve s(ub)
                            'chemin complet
                            s(ub) = sDir
                        End If
                    End If
                End If
            End If
        Loop
    End If
    'ferme l'énumèration
    FindClose hFind
End Sub

'=======================================================
'vérifie le pattern d'un fichier
'=======================================================
Private Function CheckPattern(ByVal File As String, ByVal Pattern As _
    String) As Boolean
    
Dim s As String
Dim x As Long
    
    On Error Resume Next
    
    CheckPattern = False

    'si pattern '*.*', alors tous les fichiers sont acceptés
    If Pattern = "*.*" Then
        CheckPattern = True
        Exit Function
    End If
    
    CheckPattern = CBool(InStr(1, LCase$(Pattern), _
        LCase$(GetFileExtension(File)), vbBinaryCompare))

End Function

'=======================================================
'Récupère l'extension du fichier depuis un path
'=======================================================
Public Function GetFileExtension(ByVal Path As String) As String

    GetFileExtension = String$(lstrlenA(ByVal PathFindExtension(Path)), 0)
    
    Call lstrcpyA(ByVal GetFileExtension, ByVal PathFindExtension(Path))
    
End Function
