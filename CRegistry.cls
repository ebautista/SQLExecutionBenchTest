VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
       
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
       
Private Type QueryStringResult
    StringResult As Boolean
    StringValue As String
End Type

Public Enum RegistryPathType
    cpiStandard = 1
    cpiUserDefined = 2
End Enum

Public Enum RegistryOperation
    cpiVoid = 0
    cpiGetSetting = 1
    cpiSaveSetting = 2
    cpiDeleteSetting = 3
    cpiDeleteKey = 4
    cpiEnumSettings = 5
    cpiEnumSubkeys = 6
End Enum

Public Enum TargetRegistry
    cpiCurrentUser = 1
    cpiLocalMachine = 2
    cpiCurrentConfig = 3
    cpiClassesRoot = 4
End Enum

Private strFileDSNPath As String
Private strConnectionString As String
    
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_CURRENT_CONFIG = &H80000005

Private Const REG_SZ = 1
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4

Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000

Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)


Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = (KEY_READ)
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private mvarRegistryValue As String
Private mvarRegistryOperation As RegistryOperation
Private mvarChangeOperation As Boolean
    
Public Property Get RegistryOperation() As RegistryOperation
    
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.RegistryOperation
    RegistryOperation = mvarRegistryOperation

End Property

Public Property Get RegistryValue() As String
    
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.RegistryValue
    RegistryValue = mvarRegistryValue

End Property

Function GetRegistry(ByVal Target As TargetRegistry, _
                    ByVal ApplicationName As String, _
                    ByVal RegistryKey As String, _
                    ByVal RegistrySetting As String, _
                    Optional PathType As RegistryPathType = cpiStandard, _
                    Optional ByVal DefaultValue As String = "", _
                    Optional AllowNoAdminRights As Boolean = True) _
                    As Boolean

          Dim QueryResult As QueryStringResult
          Dim lngReturnHandle As Long
          Dim strRegistryPath As String
          Dim hKey As Long
          Dim hAccess As Long
              
10        On Error GoTo ErrHandler
          
20        GetRegistry = True
30        mvarRegistryOperation = cpiGetSetting
          
40        Select Case Target
              Case cpiCurrentConfig
50                hKey = HKEY_CURRENT_CONFIG
60            Case cpiCurrentUser
70                hKey = HKEY_CURRENT_USER
80            Case cpiLocalMachine
90                hKey = HKEY_LOCAL_MACHINE
100           Case cpiClassesRoot
110               hKey = HKEY_CLASSES_ROOT
120       End Select
          
130       If PathType = cpiStandard Then
140           If IdentifyOperatingSystem = "Windows Vista" Then
150               strRegistryPath = "Software\Wow6432Node\Cubepoint\" & ApplicationName & "\" & RegistryKey
160           Else
170               strRegistryPath = "Software\Cubepoint\" & ApplicationName & "\" & RegistryKey
180           End If
190       Else
200           strRegistryPath = RegistryKey
210       End If
                      
220       hAccess = IIf(AllowNoAdminRights, KEY_READ, KEY_ALL_ACCESS)
          
          ' Open the Key
230       RegOpenKeyEx hKey, strRegistryPath, 0, hAccess, lngReturnHandle
          
          ' Get the Key's Result
240       QueryResult = RegQueryStringValue(lngReturnHandle, RegistrySetting)
250       If Not QueryResult.StringResult Then
260           GetRegistry = False
270           mvarRegistryValue = ""
280       Else
290           mvarRegistryValue = QueryResult.StringValue
              
300           If mvarRegistryValue = "" And DefaultValue <> "" Then
310               mvarChangeOperation = False
320               If Not SaveRegistry(Target, ApplicationName, RegistryKey, RegistrySetting, DefaultValue) Then
330                   GetRegistry = False
340                   Exit Function
350               End If
360           End If
370       End If
          
          ' Close the Key
380       If RegCloseKey(lngReturnHandle) <> 0 And lngReturnHandle <> 0 Then
390           GetRegistry = False
400       End If
410       Exit Function
          
ErrHandler:
420       Select Case Err.Number
              Case 0
          
430           Case Else
440               'AddToTrace "Error in GetRegistry, " & Err.Source & " (" & Err.Number & ", " & Erl & "): " & Err.Description
              
450       End Select
              
End Function

Public Function IdentifyOperatingSystem() As String
        
          Dim rOsVersionInfo As OSVERSIONINFO
          Dim sOperatingSystem As String
              
10        On Error GoTo ErrHandler
20        sOperatingSystem = "NONE"
          
          ' Pass the size of the structure into itself for the API call
          
30        rOsVersionInfo.dwOSVersionInfoSize = Len(rOsVersionInfo)
        
40        If GetVersionEx(rOsVersionInfo) Then
              
              'VISTA
              'rOsVersionInfo.dwBuildNumber : 6002
              'rOsVersionInfo.dwMajorVersion : 6
              'rOsVersionInfo.dwMinorVersion : 0
              'rOsVersionInfo.dwOSVersionInfoSize : 148
              'rOsVersionInfo.dwPlatformId : 2
              'rOsVersionInfo.szCSDVersion : Service Pack 2
              
              'WINXP
              'rOsVersionInfo.dwBuildNumber : 2600
              'rOsVersionInfo.dwMajorVersion : 5
              'rOsVersionInfo.dwMinorVersion : 1
              'rOsVersionInfo.dwOSVersionInfoSize : 148
              'rOsVersionInfo.dwPlatformId : 2
              'rOsVersionInfo.szCSDVersion : Service Pack 3
              
              'Debug.Print "rOsVersionInfo.dwBuildNumber : " & rOsVersionInfo.dwBuildNumber
              'Debug.Print "rOsVersionInfo.dwMajorVersion : " & rOsVersionInfo.dwMajorVersion
              'Debug.Print "rOsVersionInfo.dwMinorVersion : " & rOsVersionInfo.dwMinorVersion
              'Debug.Print "rOsVersionInfo.dwOSVersionInfoSize : " & rOsVersionInfo.dwOSVersionInfoSize
              'Debug.Print "rOsVersionInfo.dwPlatformId : " & rOsVersionInfo.dwPlatformId
              'Debug.Print "rOsVersionInfo.szCSDVersion : " & rOsVersionInfo.szCSDVersion
              
50            Select Case rOsVersionInfo.dwPlatformId
                  Case VER_PLATFORM_WIN32_NT
60                    If rOsVersionInfo.dwMajorVersion >= 5 Then
70                        If rOsVersionInfo.dwMajorVersion = 6 And _
                             rOsVersionInfo.dwMinorVersion = 0 Then
80                            sOperatingSystem = "Windows Vista"
90                        Else
100                           If rOsVersionInfo.dwMinorVersion = 0 Then
110                               sOperatingSystem = "Windows 2000"
120                           Else
130                               sOperatingSystem = "Windows XP"
140                           End If
150                       End If
160                   Else
170                       sOperatingSystem = "Windows NT"
180                   End If

190               Case VER_PLATFORM_WIN32_WINDOWS
200                   If rOsVersionInfo.dwMajorVersion >= 5 Then
210                       sOperatingSystem = "Windows ME"
220                   ElseIf rOsVersionInfo.dwMajorVersion = 4 And rOsVersionInfo.dwMinorVersion > 0 Then
230                       sOperatingSystem = "Windows 98"
240                   Else
250                       sOperatingSystem = "Windows 95"
260                   End If

270             Case VER_PLATFORM_WIN32s

280               sOperatingSystem = "Win32s"

290           End Select

300       End If
310     IdentifyOperatingSystem = sOperatingSystem

ErrHandler:
320       Select Case Err.Number
              Case 0
          
330           Case Else
340               'AddToTrace "Error in ConvertAccess, " & Err.Source & " (" & Err.Number & ", " & Erl & "): " & Err.Description
              
350       End Select
          

End Function




Public Function SaveRegistry(ByVal Target As TargetRegistry, _
                            ByVal ApplicationName As String, _
                            ByVal RegistryKey As String, _
                            ByVal RegistrySetting As String, _
                            ByVal NewValue As String, _
                            Optional PathType As RegistryPathType = cpiStandard) _
                            As Boolean

    Dim lngKeyHandle As Long
    Dim strRegistryPath As String
    Dim strRegistrySetting As String
    
    Dim hKey As Long
                
    
    SaveRegistry = True
    If mvarChangeOperation Then
        mvarRegistryValue = ""
        mvarRegistryOperation = cpiSaveSetting
    End If
    
    Select Case Target
        Case cpiCurrentConfig
            hKey = HKEY_CURRENT_CONFIG
        Case cpiCurrentUser
            hKey = HKEY_CURRENT_USER
        Case cpiLocalMachine
            hKey = HKEY_LOCAL_MACHINE
        Case cpiClassesRoot
            hKey = HKEY_CLASSES_ROOT
    End Select
    
    If PathType = cpiStandard Then
        strRegistryPath = "Software\Cubepoint\" & ApplicationName & "\" & RegistryKey
    Else
        strRegistryPath = RegistryKey
    End If
            
    If RegCreateKey(hKey, strRegistryPath, lngKeyHandle) <> 0 Then
        SaveRegistry = True
        
        ' To allow calling from GetRegistry
        If Not mvarChangeOperation Then
            mvarChangeOperation = True
        End If
        Exit Function
    End If
    
    If RegSetValueEx(lngKeyHandle, RegistrySetting, 0, REG_SZ, ByVal NewValue, Len(NewValue)) <> 0 Then
        SaveRegistry = False
        
        ' To allow calling from GetRegistry
        If Not mvarChangeOperation Then
            mvarChangeOperation = True
        End If
        Exit Function
    End If
    
    If RegCloseKey(lngKeyHandle) <> 0 Then
        SaveRegistry = False
        
        ' To allow calling from GetRegistry
        If Not mvarChangeOperation Then
            mvarChangeOperation = True
        End If
        Exit Function
    Else
        mvarRegistryValue = NewValue
        If Not mvarChangeOperation Then
            mvarChangeOperation = True
        End If
    End If
    
End Function

Private Function RegQueryStringValue(ByVal hKey As Long, _
                                    ByVal strValueName As String) As QueryStringResult

    Dim lngLengthMinusNUllTerminator As Long
    Dim lngDataBufferSize As Long
    Dim lngValueType As Long
    Dim lngResult As Long
    
    Dim intData As Integer
    
    Dim strBuffer As String
    
    
    RegQueryStringValue.StringResult = True
    RegQueryStringValue.StringValue = ""
    
    ' Retrieve information about the key
    lngResult = RegQueryValueEx(hKey, strValueName, 0, lngValueType, ByVal 0, lngDataBufferSize)
        
    If lngValueType = REG_SZ Then
        ' Create a buffer
        strBuffer = String(lngDataBufferSize, Chr$(0))
        
        ' Retrieve the key's content
        If RegQueryValueEx(hKey, strValueName, 0, lngValueType, ByVal strBuffer, lngDataBufferSize) <> 0 Then
            RegQueryStringValue.StringResult = False
        Else
            lngLengthMinusNUllTerminator = InStr(1, strBuffer, Chr$(0)) - 1
            If lngLengthMinusNUllTerminator > 0 Then
                RegQueryStringValue.StringValue = Left$(strBuffer, lngLengthMinusNUllTerminator)
            Else
                RegQueryStringValue.StringValue = ""
            End If
            ' RegQueryStringValue.StringValue = Left$(strBuffer, InStr(1, strBuffer, Chr$(0)) - 1)
        End If
    ElseIf lngValueType = REG_BINARY Or lngValueType = REG_DWORD Then
        
        ' Retrieve the key's value
        If RegQueryValueEx(hKey, strValueName, 0, lngValueType, intData, lngDataBufferSize) <> 0 Then
            RegQueryStringValue.StringResult = False
        Else
            RegQueryStringValue.StringValue = intData
        End If
    End If
    
End Function

Public Function DeleteRegistry(ByVal Target As TargetRegistry, _
                                ByVal ApplicationName As String, _
                                ByVal RegistryKey As String, _
                                ByVal RegistrySetting As String, _
                                Optional PathType As RegistryPathType = cpiStandard) _
                                As Boolean

    Dim lngReturnHandle As Long
    Dim strRegistryPath As String
    
    Dim hKey As Long
        
        
    mvarRegistryOperation = cpiDeleteSetting
    DeleteRegistry = True
    
    Select Case Target
        Case cpiCurrentConfig
            hKey = HKEY_CURRENT_CONFIG
        Case cpiCurrentUser
            hKey = HKEY_CURRENT_USER
        Case cpiLocalMachine
            hKey = HKEY_LOCAL_MACHINE
        Case cpiClassesRoot
            hKey = HKEY_CLASSES_ROOT
    End Select
            
    If PathType = cpiStandard Then
        strRegistryPath = "Software\Cubepoint\" & ApplicationName & "\" & RegistryKey
    Else
        strRegistryPath = RegistryKey
    End If

    ' Open the Key
    If RegOpenKeyEx(hKey, strRegistryPath, 0, KEY_ALL_ACCESS, lngReturnHandle) <> 0 Then
        DeleteRegistry = False
        mvarRegistryValue = ""
        Exit Function
    End If
        
    ' Delete the key's value
    If RegDeleteValue(lngReturnHandle, RegistrySetting) <> 0 Then
        DeleteRegistry = False
        mvarRegistryValue = ""
        Exit Function
    End If
    
    'close the key
    If RegCloseKey(lngReturnHandle) <> 0 Then
        DeleteRegistry = False
        mvarRegistryValue = ""
    Else
        mvarRegistryValue = strRegistryPath & "|||||" & RegistrySetting
    End If
    
End Function

Public Function DeleteRegistryKey(ByVal Target As TargetRegistry, _
                                    ByVal ApplicationName As String, _
                                    ByVal RegistryKey As String, _
                                    Optional PathType As RegistryPathType = cpiStandard) _
                                    As Boolean
    Dim lngReturnHandle As Long
    Dim strRegistryPath As String
    
    Dim hKey As Long
    
    
    mvarRegistryOperation = cpiDeleteKey
    DeleteRegistryKey = True
    
    Select Case Target
        Case cpiCurrentConfig
            hKey = HKEY_CURRENT_CONFIG
        Case cpiCurrentUser
            hKey = HKEY_CURRENT_USER
        Case cpiLocalMachine
            hKey = HKEY_LOCAL_MACHINE
        Case cpiClassesRoot
            hKey = HKEY_CLASSES_ROOT
    End Select
    
    If PathType = cpiStandard Then
        strRegistryPath = "Software\Cubepoint\" & ApplicationName
    Else
        strRegistryPath = RegistryKey
    End If
            
    ' Open the Key
    If RegOpenKeyEx(hKey, strRegistryPath, 0, KEY_ALL_ACCESS, lngReturnHandle) <> 0 Then
        DeleteRegistryKey = False
        mvarRegistryValue = ""
        Exit Function
    End If
        
    ' Delete the Key
    If RegDeleteKey(lngReturnHandle, RegistryKey) <> 0 Then
        DeleteRegistryKey = False
        mvarRegistryValue = ""
        Exit Function
    End If
    
    ' Close the Key
    If RegCloseKey(lngReturnHandle) <> 0 Then
        DeleteRegistryKey = False
        mvarRegistryValue = ""
    Else
        If PathType = cpiStandard Then
            mvarRegistryValue = strRegistryPath & "\" & RegistryKey
        Else
            mvarRegistryValue = RegistryKey
        End If
    End If
    
End Function

Public Function GetRegistrySubkeys(ByVal Target As TargetRegistry, _
                                    ByVal ApplicationName As String, _
                                    ByVal RootKey As String, _
                                    Optional PathType As RegistryPathType = cpiStandard) _
                                    As Boolean

    Dim lngReturnHandle As Long
    
    Dim intSubkeyCtr As Integer
    Dim hKey As Long
    
    Dim strRegistryPath As String
    Dim strSubkeyStream As String
    Dim strBufferSave As String
    
        
    mvarRegistryOperation = cpiEnumSubkeys
    
    GetRegistrySubkeys = True
    strSubkeyStream = ""
    intSubkeyCtr = 0
    
    Select Case Target
        Case cpiCurrentConfig
            hKey = HKEY_CURRENT_CONFIG
        Case cpiCurrentUser
            hKey = HKEY_CURRENT_USER
        Case cpiLocalMachine
            hKey = HKEY_LOCAL_MACHINE
        Case cpiClassesRoot
            hKey = HKEY_CLASSES_ROOT
    End Select
    
    If PathType = cpiStandard Then
        strRegistryPath = "Software\Cubepoint\" & ApplicationName & "\" & RootKey
    Else
        strRegistryPath = RootKey
    End If
        
    If RegOpenKeyEx(hKey, strRegistryPath, 0, KEY_ALL_ACCESS, lngReturnHandle) <> 0 Then
        GetRegistrySubkeys = False
        mvarRegistryValue = ""
        Exit Function
    End If
    
    Do While True
        ' Create a buffer
        strBufferSave = String(255, 0)
        
        ' Enumerate the keys
        If RegEnumKeyEx(lngReturnHandle, intSubkeyCtr, strBufferSave, 255, 0, vbNullString, ByVal 0&, ByVal 0&) <> 0 Then
            Exit Do
        End If
        
        ' Concatenate Results
        strSubkeyStream = strSubkeyStream & "|||||" & StripTerminator(strBufferSave)
        
        intSubkeyCtr = intSubkeyCtr + 1
    Loop
    
    strSubkeyStream = IIf(Trim(strSubkeyStream) <> "", Mid(strSubkeyStream, 6), strSubkeyStream)
    
    ' Close the registry key
    If RegCloseKey(lngReturnHandle) <> 0 Then
        GetRegistrySubkeys = False
        mvarRegistryValue = ""
    Else
        mvarRegistryValue = strSubkeyStream
    End If
    
End Function
Public Function GetRegistrySettings(ByVal Target As TargetRegistry, _
                                    ByVal ApplicationName As String, _
                                    ByVal RegistryKey As String, _
                                    Optional PathType As RegistryPathType = cpiStandard) _
                                    As Boolean

    Dim QueryResult As QueryStringResult
    Dim lngReturnHandle As Long
    
    Dim intSubkeyCtr As Integer
    Dim hKey As Long
    
    Dim strBufferSaveValue As String
    Dim strRegistryPath As String
    Dim strSubkeyStream As String
    Dim strBufferSave As String
    
       
    mvarRegistryOperation = cpiEnumSettings
        
    GetRegistrySettings = True
    strSubkeyStream = ""
    intSubkeyCtr = 0
    
    Select Case Target
        Case cpiCurrentConfig
            hKey = HKEY_CURRENT_CONFIG
        Case cpiCurrentUser
            hKey = HKEY_CURRENT_USER
        Case cpiLocalMachine
            hKey = HKEY_LOCAL_MACHINE
        Case cpiClassesRoot
            hKey = HKEY_CLASSES_ROOT
    End Select
    
    If PathType = cpiStandard Then
        strRegistryPath = "Software\Cubepoint\" & ApplicationName & "\" & RegistryKey
    Else
        strRegistryPath = RegistryKey
    End If
    
    If RegOpenKeyEx(hKey, strRegistryPath, 0, KEY_ALL_ACCESS, lngReturnHandle) <> 0 Then
        GetRegistrySettings = False
        mvarRegistryValue = ""
        Exit Function
    End If
        
    Do While True
        ' Create a buffer
        strBufferSave = String(255, 0)
        
        ' enumerate the values
        If RegEnumValue(lngReturnHandle, intSubkeyCtr, strBufferSave, 255, 0, ByVal 0&, ByVal 0&, ByVal 0&) <> 0 Then
            Exit Do
        End If
        
        QueryResult = RegQueryStringValue(lngReturnHandle, StripTerminator(strBufferSave))
        If Not QueryResult.StringResult Then
            GetRegistrySettings = False
            mvarRegistryValue = ""
            Exit Do
        Else
            strBufferSaveValue = QueryResult.StringValue
        
            ' Concatenate Results
            strSubkeyStream = strSubkeyStream & "|||||" & StripTerminator(strBufferSave) & "*****" & IIf(Trim(strBufferSaveValue) = "", "~~~~~", strBufferSaveValue)
            
            intSubkeyCtr = intSubkeyCtr + 1
        End If
    Loop
    
    strSubkeyStream = IIf(Trim(strSubkeyStream) <> "", Mid(strSubkeyStream, 6), strSubkeyStream)
    
    ' Close the Registry
    If RegCloseKey(lngReturnHandle) <> 0 Then
        GetRegistrySettings = False
        mvarRegistryValue = ""
    Else
        mvarRegistryValue = strSubkeyStream
    End If
    
End Function

Private Function StripTerminator(sInput As String) As String
    Dim ZeroPos As Integer
    
    'Search the first chr$(0)
    ZeroPos = InStr(1, sInput, vbNullChar)
    If ZeroPos > 0 Then
        StripTerminator = Left(sInput, ZeroPos - 1)
    Else
        StripTerminator = sInput
    End If
    
End Function

Private Sub Class_Initialize()

    mvarRegistryValue = ""
    mvarChangeOperation = True
    mvarRegistryOperation = cpiVoid
    
End Sub

Public Function RegistryKeyExists(ByVal Target As TargetRegistry, _
                                ByVal ApplicationName As String, _
                                ByVal RootKey As String, _
                                ByVal SearchKey As String) _
                                As Boolean
    
    Dim strSubkeyStream As String
    Dim strTempKey As String
    Dim strRootKey As String
    Dim strSubKey As String
    
    Dim arrSubkeys
    Dim arrValues
    
    Dim lngSubkeyCtr As Long
    Dim lngValuesCtr As Long
    Dim intCount As Long
    
    
    RegistryKeyExists = False
    strTempKey = ""
                    
    ' Get SubKeys
    If GetRegistrySettings(Target, ApplicationName, "Settings", cpiStandard) Then
        If Trim(mvarRegistryValue) <> "" Then
            arrSubkeys = Split(mvarRegistryValue, "|||||")
            For lngSubkeyCtr = 0 To UBound(arrSubkeys)
                arrValues = Split(arrSubkeys(lngSubkeyCtr), "*****")
                For lngValuesCtr = 0 To UBound(arrValues)
                    If UCase(CStr(arrValues(lngValuesCtr))) = UCase(SearchKey) Then
                        RegistryKeyExists = True
                        Exit For
                    End If
                Next
            Next
        Else
            If UCase(CStr(Trim(mvarRegistryValue))) = UCase(Trim(SearchKey)) Then
                RegistryKeyExists = True
            End If
        End If
    End If
    
    ' Make Sure it is Empty
    mvarRegistryValue = ""
    
End Function

Public Function IsAllowedAdminRights() As Boolean
    Dim QueryResult As QueryStringResult
    
    Dim lngReturnHandle As Long
    Dim hAccess As Long
    Dim hKey As Long
    
    Dim strRegistryPath As String
    
    
    IsAllowedAdminRights = True
        
    hKey = HKEY_LOCAL_MACHINE
    
    strRegistryPath = "Software\Microsoft\Windows\CurrentVersion"
    
    ' Open the Key
    RegOpenKeyEx hKey, strRegistryPath, 0, KEY_ALL_ACCESS, lngReturnHandle
    
    ' Get the Key's Result
    QueryResult = RegQueryStringValue(lngReturnHandle, "ProgramFilesDir")
    If Not QueryResult.StringResult Then
        IsAllowedAdminRights = False
    End If
    
    ' Close the Key
    If RegCloseKey(lngReturnHandle) <> 0 And lngReturnHandle <> 0 Then
        IsAllowedAdminRights = False
    End If
    
End Function

