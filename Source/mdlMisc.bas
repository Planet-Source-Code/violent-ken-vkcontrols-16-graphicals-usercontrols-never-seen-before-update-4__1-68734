Attribute VB_Name = "mdlMisc"
Option Explicit

Private Const BIF_RETURNONLYFSDIRS          As Long = 1
Private Const BIF_DONTGOBELOWDOMAIN         As Long = 2
Private Const TH32CS_SNAPPROCESS            As Long = &H2
Private Const PROCESS_TERMINAT              As Long = &H1

Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long


'=======================================================
'Affiche la boite de dialogue 'Choix du dossier'
'=======================================================
Public Function BrowseForFolder(ByVal Title As String, hwnd As Long) As String
Dim lIDList As Long
Dim sBuffer As String
Dim tBrowseInfo As BrowseInfo

    'affectation des propriétés définies par les arguments de la fonction
    With tBrowseInfo
        .hwndOwner = hwnd
        .lpszTitle = Title
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    
    'affichage de la boite de dialogue
    lIDList = SHBrowseForFolder(tBrowseInfo)
    
    'récupération du résultat
    If lIDList Then
        sBuffer = String$(260, vbNullChar)
        SHGetPathFromIDList lIDList, sBuffer
        BrowseForFolder = Mid$(sBuffer, 1, InStr(sBuffer, vbNullChar) - 1)
        CoTaskMemFree lIDList
    End If
    
End Function


'=======================================================
'tue un processus depuis son nom
'=======================================================
Public Function KickProcess(ByVal ProcessName As String) As Long
Dim hSnapshot As Long
Dim uProcess As PROCESSENTRY32
Dim lhwndProcess As Long
Dim lExitCode As Long
Dim r As Long
    
    ProcessName = LCase$(ProcessName)
    
    'création du snapshot (liste)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    
    If hSnapshot = 0 Then Exit Function
    
    'récupère la taille
    uProcess.dwSize = Len(uProcess)
    
    'premier process à être vu
    r = ProcessFirst(hSnapshot, uProcess)
    
    Do While r
        If GetFileName(uProcess.szexeFile) = ProcessName Then
        
            'ouvre le process en question
            lhwndProcess = OpenProcess(PROCESS_TERMINAT, 0, _
                uProcess.th32ProcessID)
            
            'kill le processus
            KickProcess = TerminateProcess(lhwndProcess, lExitCode)
            
            'referme le handle
            Call CloseHandle(lhwndProcess)
            
            Exit Do
        End If
        'prochain process
        r = ProcessNext(hSnapshot, uProcess)
    Loop
    
    'fermeture du handle
    Call CloseHandle(hSnapshot)
    
End Function

'=======================================================
'récupère un nom de fichier depuis un path
'=======================================================
Private Function GetFileName(ByVal sString As String) As String
Dim l As Long

    'enlève le vbnullchar de fin si nécessaire
    If InStr(sString, vbNullChar) Then sString = Left$(sString, _
        InStr(sString, vbNullChar) - 1)
    
    GetFileName = sString
End Function
