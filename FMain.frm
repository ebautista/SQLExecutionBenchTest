VERSION 5.00
Begin VB.Form FMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SQL Execution Bench Test"
   ClientHeight    =   8955
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10980
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboDtype 
      Height          =   315
      Left            =   9240
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtDetails 
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtExecutionTimes 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "&Execute"
      Height          =   495
      Left            =   8880
      TabIndex        =   4
      Top             =   8280
      Width           =   1935
   End
   Begin VB.Frame fraLog 
      Caption         =   "Execution Time Log"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   10695
      Begin VB.ListBox lstLog 
         Height          =   5715
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   10215
      End
   End
   Begin VB.ComboBox cboConnector 
      Height          =   315
      Left            =   9240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtDBPath 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
   Begin VB.Label lblDtype 
      Alignment       =   1  'Right Justify
      Caption         =   "DTYPE:"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   12
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Details:"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblExecutionTimes 
      Alignment       =   1  'Right Justify
      Caption         =   "Executions:"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblConnector 
      Alignment       =   1  'Right Justify
      Caption         =   "Connector:"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblDBPath 
      Alignment       =   1  'Right Justify
      Caption         =   "Database Path:"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************************************************************
'*
'*      Program Name:   SQL Execution Bench Test
'*      Company:        Cubepoint Inc.
'*      Author:         Edwin Bautista
'*      Started:        July 10, 2012
'*      Completed:
'*      Last Updated:
'*      Version :       1.0.0
'**************************************************************************************************************************************
'*      This utility is intended to check how long it takes to execute complicated SQL updates/inserts to benchmark the performance
'*      of DAO/ADO objects. Hopefully this test results will shed light on the problem of running Access Database on a Windows 2008
'*      Server R2 environment.
'**************************************************************************************************************************************
Option Explicit

Private m_conSadbel As ADODB.Connection
Private m_datSadbel As DAO.Database
Private m_datData As DAO.Database

Private Const CONST_ADO_CONNECTOR As String = "ADO"
Private Const CONST_DAO_CONNECTOR As String = "DAO"

Private Const IMPORT_DECLARATION As String = "Import"
Private Const COMBINED_DECLARATION As String = "Combined"

Private Const APPLICATION_NAME As String = "Clearingpoint"
Private Const DATABASE_SADBEL As String = "mdb_sadbel.mdb"
Private Const DATABASE_DATA As String = "mdb_data.mdb"


Private Sub cmdExecute_Click()
    
    cmdExecute.Enabled = False
    
    'Cleanup screen
    lstLog.Clear
    
    'Execute and Log Bench Test
    ExecuteAndLogSQLScript
    
    cmdExecute.Enabled = True
    
End Sub

Private Sub Form_Load()
    'Set Default Values
    LoadDefaultValues
    
    'Initialize Combo Box
    InitComboBox
    
    'Get MDB Path
    InitMDBPath
    
    'Initialize Connection and Database
    InitConnectionAndDatabase
End Sub


Private Sub LoadDefaultValues()
    txtDetails.Text = 10
    txtExecutionTimes.Text = 1000
End Sub


Private Sub InitComboBox()
    cboConnector.AddItem CONST_ADO_CONNECTOR
    cboConnector.AddItem CONST_DAO_CONNECTOR
    cboConnector.ListIndex = 0
    
    cboDtype.AddItem COMBINED_DECLARATION
    cboDtype.AddItem IMPORT_DECLARATION
    cboDtype.ListIndex = 0
End Sub


Private Sub InitMDBPath()
    G_strMDBPath = GetCPDatabasePath
    
    If Not FileWasFound(G_strMDBPath, DATABASE_SADBEL, Me) Then
        MsgBox DATABASE_SADBEL & " not found."
        End
    End If
    
    txtDBPath.Text = G_strMDBPath
    
End Sub


Private Sub InitConnectionAndDatabase()
    
    ConnectDB m_conSadbel, G_strMDBPath & "\" & DATABASE_SADBEL, False, G_CP_PASSWORD
    OpenDAODatabase m_datSadbel, G_strMDBPath & "\" & DATABASE_SADBEL
    OpenDAODatabase m_datData, G_strMDBPath & "\" & DATABASE_DATA
    
End Sub


Private Sub ExecuteAndLogSQLScript()
    
    Dim strUniqueCode As String
    Dim lngExec As Long
    Dim dblStart As Double, dblElapse As Double
    Dim strLog As String
    
    Dim intFreeFile As Integer
    Dim hKey As Long
    
    Dim strOSName As String
    Dim strOSVersion As String
    
    Dim arrUniqueCode() As String
    
    intFreeFile = FreeFile()
    
    If txtExecutionTimes.Text = "" Then
        MsgBox "Please indicate the number of times the SQLs are going to be executed.", vbCritical, FMain.Caption
        Exit Sub
    End If
    
    If txtDetails.Text = "" Then
        MsgBox "Please indicate the number of details for the mock declarations inserted.", vbCritical, FMain.Caption
        Exit Sub
    End If
    
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, REG_PRODUCT_KEY, 0&, KEY_QUERY_VALUE, hKey) <> ERROR_SUCCESS Then
        MsgBox "Error opening key. Will not be able to retrieve OS details.", vbCritical, FMain.Caption
        Exit Sub
    Else
        strOSName = GetRegistryValue(hKey, "ProductName")
        strOSVersion = GetRegistryValue(hKey, "CurrentVersion")
    End If
    
    If cboDtype.ListIndex = 0 Then
        If cboConnector.ListIndex = 0 Then
            Open G_strMDBPath & "\BenchTest_COMBINED_ADO_" & Format$(Now(), "YYMMDD") & "_" & Format$(Now(), "hhmmss") & ".txt" For Append As #intFreeFile
                
            lstLog.AddItem "Computer Name: " & Environ$("computername")
            lstLog.ListIndex = lstLog.NewIndex
            Print #intFreeFile, "Computer Name: " & Environ$("computername")
            
            lstLog.AddItem "OS Version: " & strOSName & " ( " & strOSVersion & " )"
            lstLog.ListIndex = lstLog.NewIndex
            Print #intFreeFile, "OS Version: " & strOSName & " ( " & strOSVersion & " )" & vbCrLf
            
            lstLog.AddItem "Executing a PLDA Combined Bench Testing with " & txtExecutionTimes.Text & " execution times for declaration with " & txtDetails.Text & " details"
            lstLog.ListIndex = lstLog.NewIndex
            Print #intFreeFile, "Executing a PLDA Combined Bench Testing with " & txtExecutionTimes.Text & " execution times for declaration with " & txtDetails.Text & " details" & vbCrLf
            
            For lngExec = 1 To txtExecutionTimes.Text
                DoEvents
                
                dblStart = Timer
                
                strUniqueCode = GenerateUniqueCode(m_datData)
                
                ExecuteADOInsertsCombined m_conSadbel, strUniqueCode
                
                'ExecuteADOUpdatesImport m_conSadbel, strUniqueCode
                
                ReDim Preserve arrUniqueCode(lngExec)
                arrUniqueCode(lngExec) = strUniqueCode
                
                dblElapse = Timer - dblStart
                
                strLog = "Executing ( " & lngExec & " of " & txtExecutionTimes.Text & " ) - Duration ( ADO ) : " & FormatNumber(dblElapse, 4, vbTrue, vbUseDefault, vbUseDefault) & " seconds"
                
                lstLog.AddItem strLog
                lstLog.ListIndex = lstLog.NewIndex
                
                Print #intFreeFile, strLog
            Next
            
            lstLog.AddItem "Cleanup Records inserted to DB..."
            lstLog.ListIndex = lstLog.NewIndex
            Print #intFreeFile, "Cleanup Records inserted to DB..."
            ExecuteCleanupADOCombined m_conSadbel, arrUniqueCode
        Else
            Open G_strMDBPath & "\BenchTest_COMBINED_DAO_" & Format$(Now(), "YYMMDD") & "_" & Format$(Now(), "hhmmss") & ".txt" For Append As #intFreeFile
            
            lstLog.AddItem "Computer Name: " & Environ$("computername")
            lstLog.ListIndex = lstLog.NewIndex
            Print #intFreeFile, "Computer Name: " & Environ$("computername")
            
            lstLog.AddItem "OS Version: " & strOSName & " ( " & strOSVersion & " )"
            lstLog.ListIndex = lstLog.NewIndex
            Print #intFreeFile, "OS Version: " & strOSName & " ( " & strOSVersion & " )" & vbCrLf
            
            lstLog.AddItem "Executing a PLDA Combined Bench Testing with " & txtExecutionTimes.Text & " execution times for declaration with " & txtDetails.Text & " details"
            lstLog.ListIndex = lstLog.NewIndex
            Print #intFreeFile, "Executing a PLDA Combined Bench Testing with " & txtExecutionTimes.Text & " execution times for declaration with " & txtDetails.Text & " details" & vbCrLf
            
            For lngExec = 1 To txtExecutionTimes.Text
                DoEvents
                
                dblStart = Timer
                
                strUniqueCode = GenerateUniqueCode(m_datData)
                
                ExecuteDAOInsertsCombined m_datSadbel, strUniqueCode
                
                'ExecuteDAOUpdatesImport m_datSadbel, strUniqueCode
                
                ReDim Preserve arrUniqueCode(lngExec)
                arrUniqueCode(lngExec) = strUniqueCode
                
                dblElapse = Timer - dblStart
                
                strLog = "Executing ( " & lngExec & " of " & txtExecutionTimes.Text & " ) - Duration ( DAO ) : " & FormatNumber(dblElapse, 4, vbTrue, vbUseDefault, vbUseDefault) & " seconds"
                
                lstLog.AddItem strLog
                lstLog.ListIndex = lstLog.NewIndex
                
                Print #intFreeFile, strLog
            Next
            
            lstLog.AddItem "Cleanup Records inserted to DB..."
            lstLog.ListIndex = lstLog.NewIndex
            Print #intFreeFile, "Cleanup Records inserted to DB..."
            ExecuteCleanupDAOCombined m_datSadbel, arrUniqueCode
        End If
    Else
        If cboConnector.ListIndex = 0 Then
            Open G_strMDBPath & "\BenchTest_IMPORT_ADO_" & Format$(Now(), "YYMMDD") & "_" & Format$(Now(), "hhmmss") & ".txt" For Append As #intFreeFile
                
            lstLog.AddItem "Computer Name: " & Environ$("computername")
            lstLog.ListIndex = lstLog.NewIndex
            Print #intFreeFile, "Computer Name: " & Environ$("computername")
            
            lstLog.AddItem "OS Version: " & strOSName & " ( " & strOSVersion & " )"
            lstLog.ListIndex = lstLog.NewIndex
            Print #intFreeFile, "OS Version: " & strOSName & " ( " & strOSVersion & " )" & vbCrLf
            
            lstLog.AddItem "Executing a PLDA Combined Bench Testing with " & txtExecutionTimes.Text & " execution times for declaration with " & txtDetails.Text & " details"
            lstLog.ListIndex = lstLog.NewIndex
            Print #intFreeFile, "Executing a PLDA Import Bench Testing with " & txtExecutionTimes.Text & " execution times for declaration with " & txtDetails.Text & " details" & vbCrLf
            
            For lngExec = 1 To txtExecutionTimes.Text
                DoEvents
                
                dblStart = Timer
                
                strUniqueCode = GenerateUniqueCode(m_datData)
                
                ExecuteADOInsertsImport m_conSadbel, strUniqueCode
                
                ExecuteADOUpdatesImport m_conSadbel, strUniqueCode
                
                ReDim Preserve arrUniqueCode(lngExec)
                arrUniqueCode(lngExec) = strUniqueCode
                
                dblElapse = Timer - dblStart
                
                strLog = "Executing ( " & lngExec & " of " & txtExecutionTimes.Text & " ) - Duration ( ADO ) : " & FormatNumber(dblElapse, 4, vbTrue, vbUseDefault, vbUseDefault) & " seconds"
                
                lstLog.AddItem strLog
                lstLog.ListIndex = lstLog.NewIndex
                
                Print #intFreeFile, strLog
            Next
            
            lstLog.AddItem "Cleanup Records inserted to DB..."
            lstLog.ListIndex = lstLog.NewIndex
            Print #intFreeFile, "Cleanup Records inserted to DB..."
            ExecuteCleanupADOImport m_conSadbel, arrUniqueCode
        Else
            Open G_strMDBPath & "\BenchTest_IMPORT_DAO_" & Format$(Now(), "YYMMDD") & "_" & Format$(Now(), "hhmmss") & ".txt" For Append As #intFreeFile
            
            lstLog.AddItem "Computer Name: " & Environ$("computername")
            lstLog.ListIndex = lstLog.NewIndex
            Print #intFreeFile, "Computer Name: " & Environ$("computername")
            
            lstLog.AddItem "OS Version: " & strOSName & " ( " & strOSVersion & " )"
            lstLog.ListIndex = lstLog.NewIndex
            Print #intFreeFile, "OS Version: " & strOSName & " ( " & strOSVersion & " )" & vbCrLf
            
            lstLog.AddItem "Executing a PLDA Combined Bench Testing with " & txtExecutionTimes.Text & " execution times for declaration with " & txtDetails.Text & " details"
            lstLog.ListIndex = lstLog.NewIndex
            Print #intFreeFile, "Executing a PLDA Import Bench Testing with " & txtExecutionTimes.Text & " execution times for declaration with " & txtDetails.Text & " details" & vbCrLf
            
            For lngExec = 1 To txtExecutionTimes.Text
                DoEvents
                
                dblStart = Timer
                
                strUniqueCode = GenerateUniqueCode(m_datData)
                
                ExecuteDAOInsertsImport m_datSadbel, strUniqueCode
                
                ExecuteDAOUpdatesImport m_datSadbel, strUniqueCode
                
                ReDim Preserve arrUniqueCode(lngExec)
                arrUniqueCode(lngExec) = strUniqueCode
                
                dblElapse = Timer - dblStart
                
                strLog = "Executing ( " & lngExec & " of " & txtExecutionTimes.Text & " ) - Duration ( DAO ) : " & FormatNumber(dblElapse, 4, vbTrue, vbUseDefault, vbUseDefault) & " seconds"
                
                lstLog.AddItem strLog
                lstLog.ListIndex = lstLog.NewIndex
                
                Print #intFreeFile, strLog
            Next
            
            lstLog.AddItem "Cleanup Records inserted to DB..."
            lstLog.ListIndex = lstLog.NewIndex
            Print #intFreeFile, "Cleanup Records inserted to DB..."
            ExecuteCleanupDAOImport m_datSadbel, arrUniqueCode
        End If
    End If
    
    lstLog.AddItem "End of SQL Bench Test..."
    lstLog.ListIndex = lstLog.NewIndex
    Print #intFreeFile, "End of SQL Bench Test..."
    Close #intFreeFile
    
    Erase arrUniqueCode
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    DisconnectDB m_conSadbel
    
    m_datData.Close
    Set m_datData = Nothing
    
    m_datSadbel.Close
    Set m_datSadbel = Nothing
    
End Sub

Private Sub txtDetails_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
        If IsNumeric(Chr(KeyAscii)) <> True Then KeyAscii = 0
    End If
End Sub


Private Sub txtExecutionTimes_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
        If IsNumeric(Chr(KeyAscii)) <> True Then KeyAscii = 0
    End If
End Sub
