Attribute VB_Name = "MTracefile"
Option Explicit
                                            
    Public Enum ERegistryGroup
        ELocalMachine = 1
        ECurrentUser = 2
    End Enum
    
    Declare Function PathAddBackslash Lib "shlwapi.dll" Alias "PathAddBackslashA" (ByVal pszPath As String) As Long
    Declare Sub PathStripPath Lib "shlwapi.dll" Alias "PathStripPathA" (ByVal pszPath As String)
    Declare Function PathStripToRoot Lib "shlwapi.dll" Alias "PathStripToRootA" (ByVal pszPath As String) As Long
    Declare Function PathIsUNC Lib "shlwapi.dll" Alias "PathIsUNCA" (ByVal pszPath As String) As Long
    Declare Function PathIsUNCServer Lib "shlwapi.dll" Alias "PathIsUNCServerA" (ByVal pszPath As String) As Long
    Declare Function PathIsUNCServerShare Lib "shlwapi.dll" Alias "PathIsUNCServerShareA" (ByVal pszPath As String) As Long
    Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
    Declare Function PathIsDirectoryEmpty Lib "shlwapi.dll" Alias "PathIsDirectoryEmptyA" (ByVal pszPath As String) As Long
    Declare Function PathIsLFNFileSpec Lib "shlwapi.dll" Alias "PathIsLFNFileSpecA" (ByVal lpName As String) As Long
    Declare Function PathIsNetworkPath Lib "shlwapi.dll" Alias "PathIsNetworkPathA" (ByVal pszPath As String) As Long
    Declare Function PathIsPrefix Lib "shlwapi.dll" Alias "PathIsPrefixA" (ByVal pszPrefix As String, ByVal pszPath As String) As Long
    Declare Function PathIsRelative Lib "shlwapi.dll" Alias "PathIsRelativeA" (ByVal pszPath As String) As Long
    Declare Function PathIsRoot Lib "shlwapi.dll" Alias "PathIsRootA" (ByVal pszPath As String) As Long
    Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

    Declare Function GetLogicalDrives Lib "kernel32" () As Long
    Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
    Declare Function IsNetworkAlive Lib "SENSAPI.DLL" (ByRef lpdwFlags As Long) As Long

    Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

    Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
    Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
    Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
    
    Public Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" _
        (lpBrowseInfo As BROWSEINFO) As Long
    Public Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" _
        (ByVal pIDList As Long, ByVal pszPath As String) As Long
    
    Public Type BROWSEINFO
        hWndOwner As Long
        pIDListRoot As Long
        pszDisplayName As String
        lpszTitle As String
        ulFlags As Long
        lpFnCallback As Long
        lParam As Long
        iImage As Long
    End Type
    
    Public Const BIF_RETURNONLYFSDIRS = &H1
    
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal length As Long)
    
    Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Public Sub AddToTrace(ByVal strTraceString As String, isADO As Boolean, timeStamp As String)
    
    Dim intFreeFile As Integer
    
    Const LOGFILE_DAO As String = "DAO_SQL_EXECUTION_BENCHTEST"
    Const LOGFILE_ADO As String = "ADO_SQL_EXECUTION_BENCHTEST"

    If G_strMDBPath = vbNullString Then Exit Sub
    
    intFreeFile = FreeFile()
    
    Open G_strMDBPath & IIf(isADO, LOGFILE_ADO, LOGFILE_DAO) & "_" & timeStamp For Output As #intFreeFile
    
    Print #intFreeFile, strTraceString
    
    Close #intFreeFile

End Sub

