VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataBaseInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: MTU                                                       **
'**                                                                        **
'** Module.....: DatabaseInterface                                         **
'**                                                                        **
'** Description: Provides database connectivity                            **
'**                                                                        **
'** History....:                                                           **
'**    03/20/02 v1.71 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2003 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit
Const CON_SQLCONNECT = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=moose1027;Initial Catalog=DVP2_R2;Data Source=(local)"
Const CON_JETCONNECT = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="

Private cnn As ADODB.Connection                         'ADO Database Connection

Public IsConnected As Boolean
Public rsProjectList As ADODB.Recordset
Public rsSettings As ADODB.Recordset
Public rsGPS As ADODB.Recordset
Public rsTxMessages As ADODB.Recordset

Public ProjectName As String
Public DebugMode As Boolean

'****************************************************************************
'**                                                                        **
'**  Procedure....:  OpenDatabase                                          **
'**                                                                        **
'**  Description..:  This routine connects to the application database.    **
'**                                                                        **
'****************************************************************************
Public Function OpenDatabase(DatabasePath As String) As Boolean
    On Error GoTo ErrorHandler
    Dim ConnectString As String, SQLcmd As String, mPC As New PerformanceCounter
    
    mPC.StartTimer True
    
    '---- Validate the Database Path - we can't run without the database!
    If UCase(Left(DatabasePath, 3)) = "SQL" Then
        '--- Need to validate SQL Connection...
        ConnectString = CON_SQLCONNECT
        AppLog InfoMsg, "OpenDatabase,Connecting to SQL Server using: " & ConnectString
    Else
        If FileSystemHandle.FileExists(DatabasePath) = False Then
            AppLog ErrorMsg, "OpenDatabase,Path [" & DatabasePath & "] to database not found."
            MsgBox "Path [" & DatabasePath & "] to database not found."
            End
        End If
        ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & ";Persist Security Info=False"
        AppLog InfoMsg, "OpenDatabase,Connecting to Micrsoft Access Database using: " & ConnectString
    End If

    AppLog InfoMsg, "OpenDatabase,Opening ADO Connection=" & ConnectString
    Set cnn = New ADODB.Connection
    cnn.CursorLocation = adUseClient
    cnn.Open ConnectString
    
    ProjectName = "Default"
    
    GetProjectListRecordset
    GetGPSRecordset
    GetTxMessagesRecordset
    
    AppLog InfoMsg, "OpenDatabase,Timed," & Format(mPC.StopTimer, "####.####") & " seconds."
    
    If FileSystemHandle.FolderExists(App.Path & "\" & DB.ProjectName) = False Then
        FileSystemHandle.CreateFolder App.Path & "\" & DB.ProjectName
    End If
    
    
    IsConnected = True
    OpenDatabase = True
    Set mPC = Nothing
    Exit Function
ErrorHandler:
    mPC.StopTimer
    Set mPC = Nothing
    AppLog ErrorMsg, "OpenDatabase, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Class_Terminate                                       **
'**                                                                        **
'**  Description..:  This routine closes recordsets & the databases.       **
'**                                                                        **
'****************************************************************************
Public Sub CloseDataBase()
    On Error GoTo 0         'ErrorHandler
    
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "Class_Terminate, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
    Resume Next
End Sub


Private Sub GetProjectListRecordset()
    On Error GoTo ErrorHandler
    Dim SQLcmd As String
    Set rsProjectList = New ADODB.Recordset
    With rsProjectList
        SQLcmd = "SELECT * FROM ProjectList ORDER BY ProjectName"
        AppLog InfoMsg, "OpenDatabase,Opening Project List Recordset as " & SQLcmd
        If .State = adStateOpen Then
            .Close
        End If
        .Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
        If .RecordCount = 0 Then
            .AddNew
            .Fields("ProjectName").Value = ProjectName
            .Fields("ProjectDescription").Value = "Default project created by ToVAAS."
            .UpdateBatch adAffectCurrent
            .Resync adAffectAllChapters
        End If
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "GetProjectListRecordset, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

Private Sub GetGPSRecordset()
    On Error GoTo ErrorHandler
    Dim SQLcmd As String
    Set rsGPS = New ADODB.Recordset
    With rsGPS
        SQLcmd = "SELECT * FROM GPS_Devices WHERE ProjectName='" & ProjectName & "' ORDER BY DeviceID"
        AppLog InfoMsg, "OpenDatabase,Opening GPS Device Recordset as " & SQLcmd
        If .State = adStateOpen Then
            .Close
        End If
        .Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "GetGPSRecordset, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

Public Function AddGPS() As Boolean
    On Error GoTo ErrorHandler
    AddGPS = False
    With rsGPS
        .AddNew
        .Fields("ProjectName").Value = ProjectName
        .Fields("DeviceID").Value = "New"
        .UpdateBatch adAffectCurrent
        .Resync adAffectAllChapters
    End With
    AddGPS = True
    Exit Function
ErrorHandler:
    AppLog ErrorMsg, "AddGPS, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Function

Private Sub GetTxMessagesRecordset()
    On Error GoTo ErrorHandler
    Dim SQLcmd As String
    Set rsTxMessages = New ADODB.Recordset
    With rsTxMessages
        SQLcmd = "SELECT * FROM VehicleMessages WHERE ProjectName='" & ProjectName & "' ORDER BY MsgID"
        AppLog InfoMsg, "OpenDatabase,Opening Transmit Vehicle Message Recordset as " & SQLcmd
        If .State = adStateOpen Then
            .Close
        End If
        .Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "GetTxMessagesRecordset, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

Public Function AddTxMessage() As Boolean
    On Error GoTo ErrorHandler
    AddTxMessage = False
    With rsTxMessages
        .AddNew
        .Fields("ProjectName").Value = ProjectName
        .Fields("MsgID").Value = "New"
        .UpdateBatch adAffectCurrent
        .Resync adAffectAllChapters
    End With
    AddTxMessage = True
    Exit Function
ErrorHandler:
    AppLog ErrorMsg, "AddTxMessage, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Function

Public Function EraseGPS() As Boolean
    On Error GoTo ErrorHandler
    EraseGPS = False
    With rsGPS
        If .RecordCount > 0 Then
            .Delete adAffectCurrent
            .Resync adAffectAllChapters
        End If
    End With
    EraseGPS = True
    Exit Function
ErrorHandler:
    AppLog ErrorMsg, "EraseGPS, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetSettingsRecordset                                  **
'**                                                                        **
'**  Description..:  This routine returns the Settings Recordset.          **
'**                                                                        **
'****************************************************************************
Private Sub GetSettingsRecordset()
    On Error GoTo ErrorHandler
    Dim SQLcmd As String
    With rsSettings
        SQLcmd = "SELECT * FROM PrinterSettings WHERE ProjectName='" & ProjectName & "' ORDER BY ProjectName,SettingNumber"
        AppLog InfoMsg, "OpenDatabase,Opening Project Settings Recordset as " & SQLcmd
        If .State = adStateOpen Then
            .Close
        End If
        .Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
    End With
    ValidateSettings
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "GetSettingsRecordset, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetServerSetting                                      **
'**                                                                        **
'**  Description..:  This routine returns settings from the database.      **
'**                                                                        **
'****************************************************************************
Public Function GetServerSetting(SettingNum As Integer) As String
    On Error GoTo ErrorHandler
    Dim SettingValue As String, SettingName As String
    With rsSettings
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "SettingNumber=" & SettingNum, 0, adSearchForward, 0
            If Not .EOF Then
                SettingValue = Trim(.Fields("SettingValue").Value & "")
                SettingName = Trim(.Fields("SettingName").Value & "")
                AppLog InfoMsg, "GetServerSetting," & SettingNum & "," & SettingName & "," & SettingValue
                GetServerSetting = SettingValue
            Else
                AppLog ErrorMsg, "GetServerSetting," & SettingNum & ",not defined."
                GetServerSetting = ""
            End If
        Else
            AppLog ErrorMsg, "GetServerSetting,No Printer Settings defined."
        End If
    End With
    Exit Function
ErrorHandler:
    AppLog ErrorMsg, "GetServerSetting,Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  CompactDatabase                                       **
'**                                                                        **
'**  Description..:  This routine compacts the print queue database.       **
'**                                                                        **
'****************************************************************************
Private Function CompactDatabase(FileName As String) As Boolean
    On Error GoTo ErrorHandler
    Dim ConnectString As String, TargetString As String, mPC As PerformanceCounter
    Set mPC = New PerformanceCounter
    mPC.StartTimer True
    AppLog InfoMsg, "CompactDatabase,Compacting " & FileName
    ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileName & ";Persist Security Info=False"
    TargetString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileName & ".tmp"
    Dim JRO As JRO.JetEngine
    Set JRO = New JRO.JetEngine
    JRO.CompactDatabase ConnectString, TargetString
    If FileSystemHandle.FileExists(FileName & ".bak") Then
        Kill FileName & ".bak"
    End If
    If FileSystemHandle.FileExists(FileName) Then
        Name FileName As FileName & ".bak"
    End If
    If FileSystemHandle.FileExists(FileName & ".tmp") Then
        Name FileName & ".tmp" As FileName
    Else
        AppLog ErrorMsg, "CompactDatabase,Error compacting print queue database - .tmp file not found."
    End If
    AppLog InfoMsg, "CompactDatabase,Timed," & Format(mPC.StopTimer, "####.####")
    Exit Function
ErrorHandler:
    mPC.StopTimer
    Set mPC = Nothing
    AppLog ErrorMsg, "CompactDatabase,Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ValidateSettings                                      **
'**                                                                        **
'**  Description..:  This routine validates printer settings in database.  **
'**                                                                        **
'****************************************************************************
Public Sub ValidateSettings()
'    CheckSetting RecCalibrationImagePath, "Calibration File Path", "C:\DVP2_R2\Calibration Images"
'    CheckSetting RecLogFilePath, "Log File Path", "C:\DVP2_R2\Log Files"
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  CheckSetting                                          **
'**                                                                        **
'**  Description..:  This routine validates printer settings in database.  **
'**                                                                        **
'****************************************************************************
Public Sub CheckSetting(SettingNum As Integer, SettingName As String, DefaultValue As String)
    On Error GoTo ErrorHandler
    With rsSettings
        If .RecordCount > 0 Then .MoveFirst
        .Find "SettingNumber=" & SettingNum
        If .EOF Then
            AppLog InfoMsg, "CheckSetting,Added Setting " & SettingNum & " Name=" & SettingName
            .AddNew
            .Fields("ProjectName").Value = ProjectName
            .Fields("SettingNumber").Value = SettingNum
            .Fields("SettingName").Value = SettingName
            .Fields("SettingValue").Value = DefaultValue
            .UpdateBatch adAffectCurrent
            .Resync adAffectAllChapters
        End If
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "CheckSetting,Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub
