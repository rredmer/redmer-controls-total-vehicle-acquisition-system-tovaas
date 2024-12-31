VERSION 5.00
Object = "{943CA7D1-C26F-4EA9-901A-2EA9BCAB0A49}#1.0#0"; "SaxComm8.ocx"
Object = "{85202277-6C76-4228-BC56-7B3E69E8D5CA}#5.0#0"; "IGToolBars50.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form frmGPSSetup 
   BorderStyle     =   0  'None
   Caption         =   "GPS Setup"
   ClientHeight    =   8040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8790
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ActiveToolBars.SSActiveToolBars GpsToolBar 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      ToolBarsCount   =   1
      ToolsCount      =   4
      Tools           =   "frmGPSSetup.frx":0000
      ToolBars        =   "frmGPSSetup.frx":32CF
   End
   Begin Threed.SSFrame DiagnosticsFrame 
      Height          =   4695
      Left            =   30
      TabIndex        =   1
      Top             =   2670
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   8281
      _Version        =   262144
      Caption         =   "Communications Diagnostics"
      Begin Threed.SSCheck ShowDiagnostics 
         Height          =   375
         Left            =   90
         TabIndex        =   3
         Top             =   270
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   661
         _Version        =   262144
         Caption         =   "Enable Diagnostics"
      End
      Begin SaxComm8Ctl.SaxComm CommGPS 
         Height          =   3945
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   690
         Visible         =   0   'False
         Width           =   8595
         _cx             =   15161
         _cy             =   6959
         Enabled         =   0   'False
         Settings        =   ",,,"
         BackColor       =   1
         Columns         =   80
         AutoProcess     =   3
         AutoScrollColumn=   -1  'True
         AutoScrollKeyboard=   -1  'True
         AutoScrollRow   =   -1  'True
         AutoSize        =   0
         BackSpace       =   0
         CaptureFilename =   ""
         CaptureMode     =   0
         CDTimeOut       =   0
         ColorFilter     =   0
         Columns         =   80
         CommEcho        =   -1  'True
         CommPort        =   "TOSHIBA Software Modem"
         CommSpy         =   0   'False
         CommSpyInput    =   0   'False
         CommSpyOutput   =   0   'False
         CommSpyProperties=   0   'False
         CommSpyWarnings =   0   'False
         CommSpyEvents   =   0   'False
         CTSTimeOut      =   0
         DialMode        =   0
         DialTimeOut     =   60000
         DSRTimeOut      =   0
         Echo            =   -1  'True
         Emulation       =   2
         EndOfLineMode   =   0
         ForeColor       =   15
         Handshaking     =   4
         IgnoreOnComm    =   0   'False
         InBufferSize    =   16384
         InputEcho       =   -1  'True
         InputLen        =   0
         InTimeOut       =   0
         OutTimeOut      =   0
         LookUpSeparator =   "|"
         LookUpText      =   ""
         LookUpTimeOut   =   10000
         NullDiscard     =   0   'False
         OutBufferSize   =   16384
         ParityReplace   =   ""
         Rows            =   500
         RThreshold      =   0
         RTSEnable       =   -1  'True
         ScrollRows      =   0
         SThreshold      =   0
         XferProtocol    =   5
         XferStatusDialog=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         StatusbarVisible=   0   'False
         ToolbarVisible  =   0   'False
         StatusDialog    =   0
         UseTAPI         =   -1  'True
         BorderStyle     =   1
         SerialNumber    =   "1180-2431098-63"
         PhoneNumber     =   ""
         ProjectFilename =   ""
         CommSpyTransfer =   0   'False
         AutoZModem      =   -1  'True
      End
      Begin SaxComm8Ctl.SaxComm CommGPS 
         Height          =   3945
         Index           =   1
         Left            =   60
         TabIndex        =   4
         Top             =   690
         Visible         =   0   'False
         Width           =   8595
         _cx             =   15161
         _cy             =   6959
         Enabled         =   0   'False
         Settings        =   ",,,"
         BackColor       =   1
         Columns         =   80
         AutoProcess     =   3
         AutoScrollColumn=   -1  'True
         AutoScrollKeyboard=   -1  'True
         AutoScrollRow   =   -1  'True
         AutoSize        =   0
         BackSpace       =   0
         CaptureFilename =   ""
         CaptureMode     =   0
         CDTimeOut       =   0
         ColorFilter     =   0
         Columns         =   80
         CommEcho        =   0   'False
         CommPort        =   "TOSHIBA Software Modem"
         CommSpy         =   0   'False
         CommSpyInput    =   0   'False
         CommSpyOutput   =   0   'False
         CommSpyProperties=   0   'False
         CommSpyWarnings =   0   'False
         CommSpyEvents   =   0   'False
         CTSTimeOut      =   0
         DialMode        =   0
         DialTimeOut     =   60000
         DSRTimeOut      =   0
         Echo            =   -1  'True
         Emulation       =   2
         EndOfLineMode   =   0
         ForeColor       =   15
         Handshaking     =   4
         IgnoreOnComm    =   0   'False
         InBufferSize    =   16384
         InputEcho       =   -1  'True
         InputLen        =   0
         InTimeOut       =   0
         OutTimeOut      =   0
         LookUpSeparator =   "|"
         LookUpText      =   ""
         LookUpTimeOut   =   10000
         NullDiscard     =   0   'False
         OutBufferSize   =   16384
         ParityReplace   =   ""
         Rows            =   25
         RThreshold      =   0
         RTSEnable       =   -1  'True
         ScrollRows      =   0
         SThreshold      =   0
         XferProtocol    =   5
         XferStatusDialog=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         StatusbarVisible=   0   'False
         ToolbarVisible  =   0   'False
         StatusDialog    =   0
         UseTAPI         =   -1  'True
         BorderStyle     =   1
         SerialNumber    =   "1180-2431098-63"
         PhoneNumber     =   ""
         ProjectFilename =   ""
         CommSpyTransfer =   0   'False
         AutoZModem      =   -1  'True
      End
      Begin SaxComm8Ctl.SaxComm CommGPS 
         Height          =   3945
         Index           =   2
         Left            =   60
         TabIndex        =   5
         Top             =   690
         Visible         =   0   'False
         Width           =   8595
         _cx             =   15161
         _cy             =   6959
         Enabled         =   0   'False
         Settings        =   ",,,"
         BackColor       =   1
         Columns         =   80
         AutoProcess     =   3
         AutoScrollColumn=   -1  'True
         AutoScrollKeyboard=   -1  'True
         AutoScrollRow   =   -1  'True
         AutoSize        =   0
         BackSpace       =   0
         CaptureFilename =   ""
         CaptureMode     =   0
         CDTimeOut       =   0
         ColorFilter     =   0
         Columns         =   80
         CommEcho        =   0   'False
         CommPort        =   "TOSHIBA Software Modem"
         CommSpy         =   0   'False
         CommSpyInput    =   0   'False
         CommSpyOutput   =   0   'False
         CommSpyProperties=   0   'False
         CommSpyWarnings =   0   'False
         CommSpyEvents   =   0   'False
         CTSTimeOut      =   0
         DialMode        =   0
         DialTimeOut     =   60000
         DSRTimeOut      =   0
         Echo            =   -1  'True
         Emulation       =   2
         EndOfLineMode   =   0
         ForeColor       =   15
         Handshaking     =   4
         IgnoreOnComm    =   0   'False
         InBufferSize    =   16384
         InputEcho       =   -1  'True
         InputLen        =   0
         InTimeOut       =   0
         OutTimeOut      =   0
         LookUpSeparator =   "|"
         LookUpText      =   ""
         LookUpTimeOut   =   10000
         NullDiscard     =   0   'False
         OutBufferSize   =   16384
         ParityReplace   =   ""
         Rows            =   25
         RThreshold      =   0
         RTSEnable       =   -1  'True
         ScrollRows      =   0
         SThreshold      =   0
         XferProtocol    =   5
         XferStatusDialog=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         StatusbarVisible=   0   'False
         ToolbarVisible  =   0   'False
         StatusDialog    =   0
         UseTAPI         =   -1  'True
         BorderStyle     =   1
         SerialNumber    =   "1180-2431098-63"
         PhoneNumber     =   ""
         ProjectFilename =   ""
         CommSpyTransfer =   0   'False
         AutoZModem      =   -1  'True
      End
      Begin SaxComm8Ctl.SaxComm CommGPS 
         Height          =   3945
         Index           =   3
         Left            =   60
         TabIndex        =   6
         Top             =   690
         Visible         =   0   'False
         Width           =   8595
         _cx             =   15161
         _cy             =   6959
         Enabled         =   0   'False
         Settings        =   ",,,"
         BackColor       =   1
         Columns         =   80
         AutoProcess     =   3
         AutoScrollColumn=   -1  'True
         AutoScrollKeyboard=   -1  'True
         AutoScrollRow   =   -1  'True
         AutoSize        =   0
         BackSpace       =   0
         CaptureFilename =   ""
         CaptureMode     =   0
         CDTimeOut       =   0
         ColorFilter     =   0
         Columns         =   80
         CommEcho        =   0   'False
         CommPort        =   "TOSHIBA Software Modem"
         CommSpy         =   0   'False
         CommSpyInput    =   0   'False
         CommSpyOutput   =   0   'False
         CommSpyProperties=   0   'False
         CommSpyWarnings =   0   'False
         CommSpyEvents   =   0   'False
         CTSTimeOut      =   0
         DialMode        =   0
         DialTimeOut     =   60000
         DSRTimeOut      =   0
         Echo            =   -1  'True
         Emulation       =   2
         EndOfLineMode   =   0
         ForeColor       =   15
         Handshaking     =   4
         IgnoreOnComm    =   0   'False
         InBufferSize    =   16384
         InputEcho       =   -1  'True
         InputLen        =   0
         InTimeOut       =   0
         OutTimeOut      =   0
         LookUpSeparator =   "|"
         LookUpText      =   ""
         LookUpTimeOut   =   10000
         NullDiscard     =   0   'False
         OutBufferSize   =   16384
         ParityReplace   =   ""
         Rows            =   25
         RThreshold      =   0
         RTSEnable       =   -1  'True
         ScrollRows      =   0
         SThreshold      =   0
         XferProtocol    =   5
         XferStatusDialog=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         StatusbarVisible=   0   'False
         ToolbarVisible  =   0   'False
         StatusDialog    =   0
         UseTAPI         =   -1  'True
         BorderStyle     =   1
         SerialNumber    =   "1180-2431098-63"
         PhoneNumber     =   ""
         ProjectFilename =   ""
         CommSpyTransfer =   0   'False
         AutoZModem      =   -1  'True
      End
      Begin SaxComm8Ctl.SaxComm CommGPS 
         Height          =   3945
         Index           =   4
         Left            =   60
         TabIndex        =   7
         Top             =   690
         Visible         =   0   'False
         Width           =   8595
         _cx             =   15161
         _cy             =   6959
         Enabled         =   0   'False
         Settings        =   ",,,"
         BackColor       =   1
         Columns         =   80
         AutoProcess     =   3
         AutoScrollColumn=   -1  'True
         AutoScrollKeyboard=   -1  'True
         AutoScrollRow   =   -1  'True
         AutoSize        =   0
         BackSpace       =   0
         CaptureFilename =   ""
         CaptureMode     =   0
         CDTimeOut       =   0
         ColorFilter     =   0
         Columns         =   80
         CommEcho        =   0   'False
         CommPort        =   "TOSHIBA Software Modem"
         CommSpy         =   0   'False
         CommSpyInput    =   0   'False
         CommSpyOutput   =   0   'False
         CommSpyProperties=   0   'False
         CommSpyWarnings =   0   'False
         CommSpyEvents   =   0   'False
         CTSTimeOut      =   0
         DialMode        =   0
         DialTimeOut     =   60000
         DSRTimeOut      =   0
         Echo            =   -1  'True
         Emulation       =   2
         EndOfLineMode   =   0
         ForeColor       =   15
         Handshaking     =   4
         IgnoreOnComm    =   0   'False
         InBufferSize    =   16384
         InputEcho       =   -1  'True
         InputLen        =   0
         InTimeOut       =   0
         OutTimeOut      =   0
         LookUpSeparator =   "|"
         LookUpText      =   ""
         LookUpTimeOut   =   10000
         NullDiscard     =   0   'False
         OutBufferSize   =   16384
         ParityReplace   =   ""
         Rows            =   25
         RThreshold      =   0
         RTSEnable       =   -1  'True
         ScrollRows      =   0
         SThreshold      =   0
         XferProtocol    =   5
         XferStatusDialog=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         StatusbarVisible=   0   'False
         ToolbarVisible  =   0   'False
         StatusDialog    =   0
         UseTAPI         =   -1  'True
         BorderStyle     =   1
         SerialNumber    =   "1180-2431098-63"
         PhoneNumber     =   ""
         ProjectFilename =   ""
         CommSpyTransfer =   0   'False
         AutoZModem      =   -1  'True
      End
   End
   Begin UltraGrid.SSUltraGrid GpsGrid 
      Height          =   2625
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   4630
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   68157460
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Override        =   "frmGPSSetup.frx":33AC
      Caption         =   "GPS Devices"
   End
End
Attribute VB_Name = "frmGPSSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: MTU                                                       **
'**                                                                        **
'** Module.....: frmGPSSetup                                               **
'**                                                                        **
'** Description: Provides GPS Configuration & Serial Communications.       **
'**                                                                        **
'** History....:                                                           **
'**    12/23/03 v1.71 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2004 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit                                     'Require explicit variable declaration
Private Const GpsLogFolder As String = "GPS Data"   'The folder name of where the GPS log files will be stored
Private Const CON_DIV As String = ","               'The field divider in the GPS log file
Public IsConnected As Boolean                       'Indicates if serial port is connected (Requires both GPS)
Public IsLogging As Boolean                         'Indicates if data is logging to files
Private GpsHandle(5) As Scripting.TextStream        'Array of GPS File Handles
Private CommBuf(5) As String                        'Array of GPS Comm Buffers
Private FileBuf(5) As String                        'Array of GPS File Buffers
Private Packets(5) As Integer                       'Number of Packets received

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Form_Load                                             **
'**                                                                        **
'**  Description..:  This routine initializes form controls.               **
'**                                                                        **
'****************************************************************************
Private Sub Form_Load()
    SetupGpsGrid                                    'Initialize form controls
    ConnectToGPS                                    'Connect to the GPS Comm Ports
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Form_Unload                                           **
'**                                                                        **
'**  Description..:  This routine closes ports & files on form exit.       **
'**                                                                        **
'****************************************************************************
Private Sub Form_Unload(Cancel As Integer)
    Dim PortNum As Integer
    IsLogging = False
    For PortNum = 0 To CommGPS.Count - 1            'Adjust each of the comm controls
        If CommGPS(PortNum).PortOpen = True Then
            CommGPS(PortNum).PortOpen = False
        End If
        If GpsHandle(PortNum) Is Nothing Then
        Else
            GpsHandle(PortNum).Close
        End If
    Next
    IsConnected = False
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Form_Resize                                           **
'**                                                                        **
'**  Description..:  This routine simply adjusts controls on form resize.  **
'**                                                                        **
'****************************************************************************
Private Sub Form_Resize()
    Dim PortNum As Integer
    If frmGPSSetup.Width = 0 Then Exit Sub
    
    GpsGrid.Width = frmGPSSetup.Width                        'Adjust Preferences grid
    DiagnosticsFrame.Width = frmGPSSetup.Width - 10          'Adjust Diagnostics frame
    For PortNum = 0 To CommGPS.Count - 1            'Adjust each of the comm controls
        CommGPS(PortNum).Width = DiagnosticsFrame.Width - 70
    Next
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GpsToolBar_ToolClick                                  **
'**                                                                        **
'**  Description..:  This routine handles the GPS Toolbar.                 **
'**                                                                        **
'****************************************************************************
Private Sub GpsToolBar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    On Error GoTo ErrorHandler
    Dim PortNum
    Select Case Tool.Id
        Case "ID_Add"
            If DB.rsGPS.RecordCount < 5 Then
                If MsgBox("Add new GPS Device?", vbApplicationModal + vbYesNo + vbQuestion + vbDefaultButton2, "Add GPS Device") = vbYes Then
                    If DB.AddGPS() = False Then
                        MsgBox "An error occured adding the new GPS device.", vbApplicationModal + vbOKOnly + vbInformation, "Error"
                    End If
                End If
            Else
                MsgBox "Only 5 GPS devices may be defined at one time.", vbApplicationModal + vbInformation + vbOKOnly, "Warning"
            End If
        Case "ID_Erase"
            If MsgBox("Erase GPS Device '" & Trim(DB.rsGPS.Fields("DeviceID").Value) & "'?", vbApplicationModal + vbYesNo + vbQuestion + vbDefaultButton2, "Add GPS Device") = vbYes Then
                If DB.EraseGPS() = False Then
                    MsgBox "An error occured erasing the new GPS device.", vbApplicationModal + vbOKOnly + vbInformation, "Error"
                End If
            End If
        Case "ID_Test"
            ConnectToGPS
        Case "ID_Start"
            If IsLogging Then
                Tool.Name = "Start"
                For PortNum = 0 To CommGPS.Count - 1
                    CommGPS(PortNum).AutoReceive = True
                    CommGPS(PortNum).Rows = 25
                Next
                IsLogging = False
            Else
                Tool.Name = "Stop"
                For PortNum = 0 To CommGPS.Count - 1
                    CommGPS(PortNum).AutoReceive = False
                    CommGPS(PortNum).Rows = 0
                Next
                IsLogging = True
            End If
    End Select
    Exit Sub
ErrorHandler:
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GpsGrid_BeforeSelectChange                            **
'**                                                                        **
'**  Description..:  This routine updates the diagnostics on record change.**
'**                                                                        **
'****************************************************************************
Private Sub GpsGrid_BeforeSelectChange(ByVal SelectChange As UltraGrid.Constants_SelectChange, ByVal NewSelections As UltraGrid.SSSelected, ByVal Cancel As UltraGrid.SSReturnBoolean)
    '---- Handle selection of new GPS Row
    If ShowDiagnostics.Value = -1 Then
        Dim Tmp As Integer, PortNum As Integer
        PortNum = DB.rsGPS.AbsolutePosition - 1
        For Tmp = 0 To CommGPS.Count - 1
            CommGPS(Tmp).Enabled = False
            CommGPS(Tmp).Visible = False
        Next
        DiagnosticsFrame.Caption = "Communication Diagnostics - " & DB.rsGPS.Fields("DeviceID").Value
        CommGPS(PortNum).Enabled = True
        CommGPS(PortNum).Visible = True
    End If
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ShowDiagnostics_Click                                 **
'**                                                                        **
'**  Description..:  This routine toggles the communications diagnostics.  **
'**                                                                        **
'****************************************************************************
Private Sub ShowDiagnostics_Click(Value As Integer)
    Dim PortNum As Integer
    PortNum = DB.rsGPS.AbsolutePosition - 1
    If Value = True Then
        DiagnosticsFrame.Caption = "Communication Diagnostics - " & DB.rsGPS.Fields("DeviceID").Value
        CommGPS(PortNum).Enabled = True
        CommGPS(PortNum).Visible = True
    Else
        DiagnosticsFrame.Caption = "Communication Diagnostics"
        CommGPS(PortNum).Enabled = False
        CommGPS(PortNum).Visible = False
    End If
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  CommGPS_Receive                                       **
'**                                                                        **
'**  Description..:  This routine handles in-bound communications.         **
'**                                                                        **
'****************************************************************************
Private Sub CommGPS_Receive(Index As Integer)
    '---- Receive function
    Dim Msgs() As String, MsgNum As Integer
    If IsLogging Then
        CommBuf(Index) = CommBuf(Index) & CommGPS(Index).Input
        Msgs = Split(CommBuf(Index), vbCrLf)
        If UBound(Msgs) > 0 Then
            For MsgNum = 0 To UBound(Msgs) - 1
                If Left(Msgs(MsgNum), 1) = "$" Then
                    FileBuf(Index) = FileBuf(Index) & Format$(Now, "mm/dd/yyyy") & CON_DIV & Format$(Now, "hh:mm:ssAMPM") & CON_DIV & Format(Timer / 1000, "00000000.000") & CON_DIV & Msgs(MsgNum) & vbCrLf
                    Packets(Index) = Packets(Index) + 1
                    'Debug.Print Packets(Index)
                    If Packets(Index) > 512 Then
                        GpsHandle(Index).Write FileBuf(Index)
                        FileBuf(Index) = ""
                        Packets(Index) = 0
                    End If
                Else
                    
                End If
            Next
            CommBuf(Index) = Msgs(MsgNum)
        End If
    End If
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  SetupGpsGrid                                          **
'**                                                                        **
'**  Description..:  This routine initializes the GPS Preferences Grid.    **
'**                                                                        **
'****************************************************************************
Private Sub SetupGpsGrid()
    
    '---- Make sure the GPS Logging folder exists
    If FileSystemHandle.FolderExists(App.Path & "\" & DB.ProjectName & "\" & GpsLogFolder) = False Then
        FileSystemHandle.CreateFolder App.Path & "\" & DB.ProjectName & "\" & GpsLogFolder
    End If
    
    '---- Initialize Comm Controls
    Dim PortNum As Integer
    For PortNum = 0 To CommGPS.Count - 1
        CommGPS(PortNum).Enabled = False
        CommGPS(PortNum).IgnoreOnComm = True
        CommGPS(PortNum).Visible = False
        CommGPS(PortNum).AutoReceive = False
        CommGPS(PortNum).AutoSend = False
        CommGPS(PortNum).Echo = False
        CommGPS(PortNum).InputEcho = True
        CommGPS(PortNum).RThreshold = 1024
        CommGPS(PortNum).InTimeOut = 1
        CommBuf(PortNum) = ""
        FileBuf(PortNum) = ""
        Packets(PortNum) = 0
    Next
    
    '---- Initialize the GPS Grid
    With GpsGrid
        Dim CommPortTemp As Variant
        
        Set .DataSource = DB.rsGPS
        .Refresh ssRefetchAndFireInitializeRow
        
        '--- Define Value Lists
        .ValueLists.Add "GPSBrands"
        .ValueLists.Item("GPSBrands").ValueListItems.Add 1, "Javad"
        
        .ValueLists.Add "GPSModels"
        .ValueLists.Item("GPSModels").ValueListItems.Add 1, "Legacy-E"
        
        .ValueLists.Add "CommPorts"
        For Each CommPortTemp In CommGPS(0).CommPorts
            .ValueLists.Item("CommPorts").ValueListItems.Add CommPortTemp
        Next
        
        .ValueLists.Add "BaudRates"
        .ValueLists.Item("BaudRates").ValueListItems.Add 0, "9600"
        .ValueLists.Item("BaudRates").ValueListItems.Add 1, "19200"
        .ValueLists.Item("BaudRates").ValueListItems.Add 2, "34800"
        .ValueLists.Item("BaudRates").ValueListItems.Add 3, "57600"
        .ValueLists.Item("BaudRates").ValueListItems.Add 4, "115200"
        .ValueLists.Item("BaudRates").ValueListItems.Add 5, "230400"
        .ValueLists.Item("BaudRates").ValueListItems.Add 6, "460800"
        .ValueLists.Item("BaudRates").ValueListItems.Add 7, "921600"
        
        .ValueLists.Add "DataBits"
        .ValueLists.Item("DataBits").ValueListItems.Add 0, "5"
        .ValueLists.Item("DataBits").ValueListItems.Add 1, "6"
        .ValueLists.Item("DataBits").ValueListItems.Add 2, "7"
        .ValueLists.Item("DataBits").ValueListItems.Add 3, "8"
        
        .ValueLists.Add "StopBits"
        .ValueLists.Item("StopBits").ValueListItems.Add 0, "1"
        .ValueLists.Item("StopBits").ValueListItems.Add 1, "1.5"
        .ValueLists.Item("StopBits").ValueListItems.Add 2, "2"
        
        .ValueLists.Add "Parity"
        .ValueLists.Item("Parity").ValueListItems.Add 0, "None"
        .ValueLists.Item("Parity").ValueListItems.Add 1, "Even"
        .ValueLists.Item("Parity").ValueListItems.Add 2, "Odd"
        .ValueLists.Item("Parity").ValueListItems.Add 3, "Mark"
        .ValueLists.Item("Parity").ValueListItems.Add 4, "Space"
        
        .ValueLists.Add "Handshaking"
        .ValueLists.Item("Handshaking").ValueListItems.Add 1, "None"
        .ValueLists.Item("Handshaking").ValueListItems.Add 2, "XonXoff"
        .ValueLists.Item("Handshaking").ValueListItems.Add 3, "Hardware"
        .ValueLists.Item("Handshaking").ValueListItems.Add 4, "Both"
        .ValueLists.Item("Handshaking").ValueListItems.Add 5, "Default"
        
        .ValueLists.Add "Status"
        .ValueLists.Item("Status").ValueListItems.Add 0, "Closed"
        .ValueLists.Item("Status").ValueListItems.Add 1, "Active"
        .ValueLists.Item("Status").ValueListItems.Add 2, "Error"
        
        .Bands(0).Columns(0).Hidden = True             'Hide project name column
        
        .Bands(0).Columns(1).Header.Caption = "GPS ID"
        .Bands(0).Columns(1).Width = 800
        
        .Bands(0).Columns(2).Style = ssStyleCheckBox
        
        .Bands(0).Columns(3).Style = ssStyleDropDownList
        .Bands(0).Columns(3).ValueList = "GPSBrands"
        
        .Bands(0).Columns(4).Style = ssStyleDropDownList
        .Bands(0).Columns(4).ValueList = "GPSModels"
        
        .Bands(0).Columns(5).Style = ssStyleDropDownList
        .Bands(0).Columns(5).ValueList = "CommPorts"
        
        .Bands(0).Columns(6).Style = ssStyleDropDownList
        .Bands(0).Columns(6).ValueList = "Baudrates"
        
        .Bands(0).Columns(7).Style = ssStyleDropDownList
        .Bands(0).Columns(7).ValueList = "DataBits"
        
        .Bands(0).Columns(8).Style = ssStyleDropDownList
        .Bands(0).Columns(8).ValueList = "StopBits"
        
        .Bands(0).Columns(9).Style = ssStyleDropDownList
        .Bands(0).Columns(9).ValueList = "Parity"
        
        .Bands(0).Columns(10).Style = ssStyleDropDownList
        .Bands(0).Columns(10).ValueList = "Handshaking"
        
        .Bands(0).Columns(11).Style = ssStyleDropDownList
        .Bands(0).Columns(11).ValueList = "Status"
    End With
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ConnectToGPS                                          **
'**                                                                        **
'**  Description..:  This routine connects to the GPS Comm Ports.          **
'**                                                                        **
'****************************************************************************
Public Sub ConnectToGPS()
    On Error GoTo ErrorHandler
    Dim PortNum As Integer, PortSettings As String
    
    IsConnected = False                             'Set connected flag to false - this simply provides status on main screen
    IsLogging = False                               'Set logging flag to flase - can't log data until connected
    
    '---- All GPS Settings are stored in the GPS Recordset which is provided by the database class
    AppLog InfoMsg, "ConnectToGPS,Connecting to GPS devices..."
    With DB.rsGPS
        If .RecordCount > 0 Then                    'If there are records for GPS devices in the database
            PortNum = 0                             'Start initializing at Port 0 (this is the comm control array on the form)
            .MoveFirst                              'Go to the first record
            Do While Not .EOF                       'Loop for each record in the set
                '---- If the device enabled field is set true then configure port settings out the database
                If .Fields("DeviceEnabled").Value = True Then
                    PortSettings = GpsGrid.ValueLists("Baudrates").Find(.Fields("BaudRate").Value, ssValueListFindDataValue, 0).DisplayText
                    PortSettings = PortSettings & "," & LCase(Left(GpsGrid.ValueLists("Parity").Find(.Fields("Parity").Value, ssValueListFindDataValue, 0).DisplayText, 1))
                    PortSettings = PortSettings & "," & GpsGrid.ValueLists("DataBits").Find(.Fields("DataBits").Value, ssValueListFindDataValue, 0).DisplayText
                    PortSettings = PortSettings & "," & GpsGrid.ValueLists("StopBits").Find(.Fields("StopBits").Value, ssValueListFindDataValue, 0).DisplayText
                    CommGPS(PortNum).CommPort = .Fields("CommPort").Value
                    CommGPS(PortNum).Settings = PortSettings
                    CommGPS(PortNum).Handshaking = .Fields("HandShaking").Value
                    AppLog InfoMsg, "ConnectToGPS,Opening GPS='" & Trim(.Fields("DeviceID").Value) & "',Port='" & CommGPS(PortNum).CommPort & "',Settings='" & CommGPS(PortNum).Settings & "',Handshaking=" & CommGPS(PortNum).Handshaking
                    
                    '---- If the serial port is already open, then close it
                    If CommGPS(PortNum).PortOpen = True Then
                        CommGPS(PortNum).PortOpen = False
                    End If
                    
                    '---- Open the serial port
                    CommGPS(PortNum).PortOpen = True
                    
                    '---- Check to see that the port opened properly
                    If CommGPS(PortNum).PortOpen = False Then
                        .Fields("Status").Value = 2
                        AppLog InfoMsg, "ConnectToGPS,Failed To Open GPS='" & Trim(.Fields("DeviceID").Value) & "',Port='" & CommGPS(PortNum).CommPort & "',Settings='" & CommGPS(PortNum).Settings & "',Handshaking=" & CommGPS(PortNum).Handshaking
                    Else
                        .Fields("Status").Value = 1
                        Set GpsHandle(PortNum) = FileSystemHandle.CreateTextFile(App.Path & "\" & DB.ProjectName & "\" & GpsLogFolder & "\" & Trim(.Fields("DeviceID").Value) & ".txt", True, False)
                        IsConnected = True
                    End If
                    .UpdateBatch adAffectCurrent
                End If
                PortNum = PortNum + 1
                .MoveNext
            Loop
            '---- Go to the first record to prevent errors in user interface handling (do not leave at EOF)
            .MoveFirst
        End If
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "ConnectToGPS,Error connecting to Comm Port."
    Resume Next
End Sub
