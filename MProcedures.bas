Attribute VB_Name = "MProcedures"
Option Explicit

Public Const ERROR_SUCCESS = 0&
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const KEY_QUERY_VALUE = &H1
Public Const REG_PRODUCT_KEY As String = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"

Public G_strMDBPath As String

Public Const G_CP_Template = "TemplateCP.mdb"
Public Const G_CONST_DB_PWORD_NONE = ""
Public Const G_CP_PASSWORD = "wack2"
Public Const CONST_DB_NAME_HISTORY = "mdb_history*.mdb"
Public Const CONST_DB_NAME_REPERTORY = "mdb_repertory*.mdb"

Public G_conTemplateCP As ADODB.Connection

Private Enum AccessVersion
    ACCESS_97 = 1997
    ACCESS_2000 = 2000
    ACCESS_2002_2003 = 2002
End Enum

Public Enum MessagingTypeConstant
    [EDI Messaging] = 1
    [XML Messaging] = 2
    [EDI Follow-up Request Messaging] = 3 'CSCLP-578
    [Unknown] = 99
End Enum

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long '1 = Windows 95.
    '2 = Windows NT
    szCSDVersion As String * 128
End Type

Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long


Public Function FindVersion(ByVal strDb As String, ByVal Password As String) As Long
    
    Dim dbs As DAO.Database
    Dim strVersion As String
    
    On Error GoTo Err_FindVersion
    
    'Open the database and return a reference to it.
    Set dbs = OpenDatabase(Name:=G_strMDBPath & "\" & strDb, _
                           Options:=False, _
                           ReadOnly:=False, _
                           Connect:=";pwd=" & Password)
    
    'Check the value of the AccessVersion property.
    strVersion = dbs.Properties("AccessVersion")
    
   'Return the two leftmost digits of the value of
    'the AccessVersion property.
    strVersion = Left(strVersion, 2)
    
    'Based on the value of the AccessVersion property,
    'return a long indicating the version of Microsoft Access
    'used to create or open the database.
    Select Case strVersion
        Case "07", "3."
            FindVersion = AccessVersion.ACCESS_97
        
        Case "08"
            FindVersion = AccessVersion.ACCESS_2000
        
        Case "09"
            FindVersion = AccessVersion.ACCESS_2002_2003
            
        Case Else
            FindVersion = 0
            
    End Select
    
    On Error Resume Next
    
    dbs.Close
    Set dbs = Nothing
    
    Exit Function
    
Err_FindVersion:
    If Err.Number = 3270 Then
        strVersion = dbs.Properties("Version")
        Resume Next
    Else
        MsgBox "Error: " & Err & vbCrLf & Err.Description
        FindVersion = 0
    End If
End Function

Public Function GetCPDatabasePath() As String
    Dim strCPDatabasePath As String
    Dim clsRegistry As CRegistry
    
    Set clsRegistry = New CRegistry
    
    clsRegistry.GetRegistry cpiLocalMachine, "Clearingpoint", "Settings", "mdbPath"
    
    If Len(Trim(clsRegistry.RegistryValue)) > 0 Then
        strCPDatabasePath = clsRegistry.RegistryValue
    Else
        strCPDatabasePath = App.Path
    End If

    Set clsRegistry = Nothing
    
    GetCPDatabasePath = strCPDatabasePath
End Function
Public Sub OpenDAODatabase(ByRef DAODatabase As DAO.Database, ByVal DBPathAndName As String)
    
    If (DAODatabase Is Nothing = False) Then
        Set DAODatabase = Nothing
    End If
    
    Set DAODatabase = OpenDatabase(Name:=DBPathAndName, _
                                   Options:=False, _
                                   ReadOnly:=False, _
                                   Connect:=";pwd=" & G_CP_PASSWORD)
    
End Sub


Public Sub ConnectDB(ByRef ADOConnection As ADODB.Connection, DatabasePath As String, UseDataShaping As Boolean, Optional ByVal Password As String)
    If Not ADOConnection Is Nothing Then
        If ADOConnection.State = adStateOpen Then
            ADOConnection.Close
        End If
        Set ADOConnection = Nothing
    End If
    Set ADOConnection = New ADODB.Connection
    
    If UseDataShaping Then
        
        ADOConnection.Provider = "MSDataShape"
        
        If Len(Trim(Password)) > 0 Then
            ADOConnection.Open "Data Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & ";Persist Security Info=False;Jet OLEDB:Database Password=" & Password
        Else
            ADOConnection.Open "Data Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & ";Persist Security Info=False"
        End If
        
    Else
        If Len(Trim(Password)) > 0 Then
            ADOConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & ";Persist Security Info=False;Jet OLEDB:Database Password=" & Password
        Else
            ADOConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & ";Persist Security Info=False"
        End If
    End If
End Sub

Public Sub DisconnectDB(ByRef conToDisconnect As ADODB.Connection)
    If Not conToDisconnect Is Nothing Then
        If conToDisconnect.State = adStateOpen Then
            conToDisconnect.Close
        End If
        Set conToDisconnect = Nothing
    End If
End Sub

Public Sub RstOpen(Source As String, ByRef conToUse As ADODB.Connection, rstToOpen As ADODB.Recordset, CursorType As CursorTypeEnum, LockType As LockTypeEnum, Optional lngCacheSize As Long = 1, Optional ByVal MakeOffline As Boolean = False)
    On Error GoTo ERROR_HANDLER_BOOKMARK
    
START:
    If Not rstToOpen Is Nothing Then
        If rstToOpen.State = adStateOpen Then
            rstToOpen.Close
        End If
        Set rstToOpen = Nothing
    End If
    Set rstToOpen = New ADODB.Recordset
    If MakeOffline = True Then
        rstToOpen.CursorLocation = adUseClient
    End If
    
    rstToOpen.CacheSize = lngCacheSize
        
    'Debug.Print Source
    
    
    rstToOpen.Open Source, conToUse, CursorType, LockType
    On Error GoTo 0
    
    If MakeOffline = True Then
        Set rstToOpen.ActiveConnection = Nothing
    End If
    
    On Error GoTo 0
    
    Exit Sub
    
ERROR_HANDLER_BOOKMARK:
    Select Case Err.Number
        Case -2147467259
            Err.Clear
            Set rstToOpen = Nothing
            GoTo START
        Case Else
            Err.Raise Err.Number, , Err.Description
    End Select
End Sub

Public Sub RstClose(rstToClose As ADODB.Recordset)

    If Not rstToClose Is Nothing Then
        If rstToClose.State = adStateOpen Then
            rstToClose.Close
        End If
        Set rstToClose = Nothing
    End If

End Sub

Public Function SetDatabasePassword(ByVal DatabasePathName As String, _
                                ByVal OldDatabasePassword As String, _
                                ByVal NewDatabasePassword As String) As Boolean
    
    Dim dbCP As DAO.Database
    
    On Error GoTo ExitFunc
    
    Set dbCP = OpenDatabase(Name:=DatabasePathName, _
                        Options:=True, _
                        ReadOnly:=False, _
                        Connect:=";pwd=" & OldDatabasePassword)

    If Not dbCP Is Nothing Then
        dbCP.NewPassword OldDatabasePassword, NewDatabasePassword
        dbCP.Close
    End If
    
    Set dbCP = Nothing

ExitFunc:
    'if Password is already present exit function
    If Err.Number = 3031 Then
        Exit Function
    End If
    
End Function

Public Function FileWasFound(FilePath As String, FileName As String, ByRef CallingForm As FMain) As Boolean

    On Error Resume Next
    If Len(Dir(FilePath & "\" & FileName)) > 0 Then
        FileWasFound = True
    Else
        If Err.Number <> 0 Then
            MsgBox "Cannot access " & FilePath & "\" & FileName & "." & _
                                vbCrLf & Err.Number & " : " & Err.Description, vbInformation, "Check File - " & Erl, Err.HelpFile, Err.HelpContext
            Unload CallingForm
            End
        End If
        On Error GoTo 0
    End If
End Function


Public Function GetRegistryValue(ByVal hKey As Long, ByVal subkey_name As String) As String
    Dim value As String
    Dim length As Long
    Dim value_type As Long
    
    On Error GoTo ErrHandler
    
    length = 256
    value = Space$(length)
    If RegQueryValueEx(hKey, subkey_name, 0&, value_type, ByVal value, length) <> ERROR_SUCCESS Then
        value = "<Error>"
    Else
        ' Remove the trailing null character.
        value = Left$(value, length - 1)
    End If
    
    GetRegistryValue = value

ErrHandler:
    Select Case Err.Number
        Case 0
            'Do Nothing
        
        Case Else
            MsgBox "Error in GetRegistryValue: " & Err.Source & " (" & Err.Number & ", " & Erl & "): " & Err.Description, vbCritical, FMain.Caption
            End
            
    End Select
    
End Function


Public Function IsIndexExisting(ByRef datSadbel As DAO.Database, ByVal IndexName As String, ByVal TableName As String) As Boolean
    Dim strObjectName As String
    Dim idxTarget As DAO.Index
    Dim TargetTable As DAO.TableDef
    
    IsIndexExisting = False
    
    datSadbel.TableDefs.Refresh
    
    For Each TargetTable In datSadbel.TableDefs
        If TargetTable.Name = TableName Then
            For Each idxTarget In TargetTable.Indexes
                If UCase(idxTarget.Name) = UCase(IndexName) Then
                    IsIndexExisting = True
                    Exit Function
                End If
            Next
            
            Exit Function
        End If
    Next
    
End Function

