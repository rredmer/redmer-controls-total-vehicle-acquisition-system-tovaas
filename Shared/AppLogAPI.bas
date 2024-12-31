Attribute VB_Name = "AppLogAPI"
Option Explicit
Public FileSystemHandle As New Scripting.FileSystemObject
Public IsLogFileConnected As Boolean                'Set TRUE when Log File is functional
Private Const CON_DIV = ":"                         'Log sectional divider
Private Const CON_LOGFILEDIR = "\LogFiles"          'Log file directory
Private Const CON_MAXFILESIZE = 1400000             'Maximum log file size
Private oFile As Scripting.TextStream               'Pointer to Log File
Private sLogDir As String                           'Pointer to log file directory
Private sLogFileName As String                      'The name of the current log file
Private lMaxLogFileSize As Long                     'Maximum log file size
Private bAutoCleanup As Boolean                     'Indicates automatic log file cleanup
Private LogFileSize As Long
Enum LogMsgTypes                                    'This enumeration helps programming
    InfoMsg = 0
    DebugMsg = 1
    ErrorMsg = 2
End Enum

'****************************************************************************
'**                                                                        **
'** Subroutine.: InitLogFile                                               **
'**                                                                        **
'** Description: This routine initializes the log file                     **
'**                                                                        **
'**              Log File Format:                                          **
'**                System Date: 10 Bytes, mm/dd/yyyy                       **
'**                System Time: 8 bytes, hh:mm:ss                          **
'**                Elapsed Milliseconds: 12 Bytes, ssssssss.mmm            **
'**                Message Type: 5 bytes, enumerated constants (see below).**
'**                Message Text: Message text from calling function.       **
'**                                                                        **
'**              Message Types:                                            **
'**                INFOR: General information messages - always logged.    **
'**                DEBUG: Application debugging enabled, log for debug only**
'**                ERROR: Application error - always logged.               **
'**                                                                        **
'**              CONFIGURATION:                                            **
'**                DebugEnable - Set 1 to log DEBUG messages, 0 to ignore. **
'**                AutoCleanup - Set 1 to auto delete log files > 1 month. **
'**                MaxLogSize  - Determines maximum log file size.         **
'**                * Configurable values are stored in the Registry.       **
'**                                                                        **
'****************************************************************************
Public Sub InitLogFile()
    On Error GoTo ErrorHandler
    lMaxLogFileSize = CON_MAXFILESIZE
    bAutoCleanup = True
    Set FileSystemHandle = New Scripting.FileSystemObject
    LogFileConnect
    Exit Sub
ErrorHandler:
End Sub

Public Sub CloseLogFile()
    LogFileDisconnect
    Set FileSystemHandle = Nothing
End Sub

'****************************************************************************
'**                                                                        **
'** Subroutine.: LogFileConnect                                            **
'**                                                                        **
'** Description: This routine creates a new log file instance.             **
'**                                                                        **
'****************************************************************************
Private Sub LogFileConnect()
    On Error GoTo ErrorHandler
    IsLogFileConnected = False
    '---- Get pointer to filesystem object and verify log file folder
    sLogDir = App.Path & CON_LOGFILEDIR
    If FileSystemHandle.FolderExists(sLogDir) = False Then
        FileSystemHandle.CreateFolder sLogDir
    End If
    '---- Get pointer to current log file and create it if necessary
    sLogFileName = sLogDir & "\LOG_" & Format$(Now, "mmddyyyy") & "_" & Format$(Now, "HhNnSs") & ".TXT"
    Set oFile = FileSystemHandle.CreateTextFile(sLogFileName, True, False)
    LogFileSize = 0
    'AppLog InfoMsg, "Created log file (" & sLogFileName & ")"
    IsLogFileConnected = True
    Exit Sub
ErrorHandler:
End Sub

'****************************************************************************
'**                                                                        **
'** Subroutine.: LogFileDisconnect                                         **
'**                                                                        **
'** Description: This routine closes the currently open log file.          **
'**                                                                        **
'****************************************************************************
Private Sub LogFileDisconnect()
    On Error GoTo ErrorHandler
    oFile.Close
    Set oFile = Nothing
    IsLogFileConnected = False
    Exit Sub
ErrorHandler:
End Sub

'****************************************************************************
'**                                                                        **
'** Subroutine.: Log                                                       **
'**                                                                        **
'** Description: This routine writes a new line of text to the log file.   **
'**                                                                        **
'****************************************************************************
Public Sub AppLog(lMsgType As LogMsgTypes, sLogMessage As String)
    '---- If debug message sent and debugging disabled, exit sub
    If DB.DebugMode = False And lMsgType = DebugMsg Then
        Exit Sub
    End If
    If IsLogFileConnected = False Then
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    '---- Choose message type
    Dim sMessageType As String, sMsg As String
    Select Case lMsgType
        Case 0
            sMessageType = "INFOR"
        Case 1
            sMessageType = "DEBUG"
        Case 2
            sMessageType = "ERROR"
            
            MsgBox sLogMessage
            
    End Select
    '---- Write message into log file
    sMsg = Format$(Now, "mm/dd/yyyy") & CON_DIV & Format$(Now, "hh:mm:ssAMPM") & CON_DIV & Format(Timer / 1000, "00000000.000") & CON_DIV & sMessageType & CON_DIV & Trim$(sLogMessage)
    oFile.WriteLine sMsg
    LogFileSize = LogFileSize + Len(sMsg)
    '---- Create new log file when current file reaches 1.4MB (this provides for copying files to Floppy disk)
    If LogFileSize > lMaxLogFileSize Then
        LogFileDisconnect
        LogFileConnect
    End If
    Exit Sub
ErrorHandler:
End Sub

'****************************************************************************
'**                                                                        **
'** Subroutine.: DeleteLogFiles                                            **
'**                                                                        **
'** Description: This routine deletes log files, either forcibly through   **
'**              the user interface or programmatically using AutoCleanup. **
'**                                                                        **
'** Returns....: The number of files deleted.                              **
'**                                                                        **
'****************************************************************************
Private Function DeleteLogFiles(bCalledByAutoCleanup As Boolean) As Long
    On Error GoTo ErrorHandler
    Dim lFileCount As Long
    Dim oFolder As Scripting.Folder
    Dim oFl As Scripting.File
    Set oFolder = FileSystemHandle.GetFolder(sLogDir)
    lFileCount = 0
    For Each oFl In oFolder.Files
        If UCase$(Left$(oFl.Name, 3)) = "LOG" And UCase$(Right$(oFl.Name, 3)) = "TXT" Then
            If bCalledByAutoCleanup Then
                '---- Delete all log files not created in current month
                If Month(Now) <> Month(oFl.DateLastModified) Then
                    oFl.Delete True
                End If
            Else
                oFl.Delete True
                lFileCount = lFileCount + 1
            End If
        End If
    Next
    Set oFl = Nothing
    Set oFolder = Nothing
    DeleteLogFiles = lFileCount
    Exit Function
ErrorHandler:
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  MakeCstring                                           **
'**                                                                        **
'**  Description..:  This routine converts VB file string to C format.     **
'**                                                                        **
'****************************************************************************
Public Function MakeCstring(SourceString As String) As String
    Dim ByteNum As Integer, Target As String
    For ByteNum = 1 To Len(SourceString)
        If Mid(SourceString, ByteNum, 1) = "\" Then
            Target = Target & "\\"
        Else
            Target = Target & Mid(SourceString, ByteNum, 1)
        End If
    Next
    MakeCstring = Target
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  MyHex                                                 **
'**                                                                        **
'**  Description..:  This routine converts Decimal to formatted Hex.       **
'**                                                                        **
'****************************************************************************
Public Function MyHex(DecValue As Currency, StringLength As Integer) As String
    Dim ByteNum As Integer, HexString As String
    HexString = Hex(DecValue)
    For ByteNum = Len(HexString) To StringLength - 1
        HexString = "0" & HexString
    Next
    MyHex = HexString
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  DecimalToBinary                                       **
'**                                                                        **
'**  Description..:  This routine converts Decimal to Binary.              **
'**                                                                        **
'****************************************************************************
Public Function DecimalToBinary(ByVal lngDecimal As Long) As Byte()
    Dim intIndex As Integer, bytBinaryNumber(31) As Byte
    For intIndex = 0 To 31
        If lngDecimal Mod 2 = 1 Then
            bytBinaryNumber(intIndex) = 1
        Else
            bytBinaryNumber(intIndex) = 0
        End If
        lngDecimal = lngDecimal \ 2
        If lngDecimal = 0 Then
            Exit For
        End If
    Next
    DecimalToBinary = bytBinaryNumber
End Function

