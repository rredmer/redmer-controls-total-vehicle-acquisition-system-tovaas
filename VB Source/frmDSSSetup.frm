VERSION 5.00
Object = "{D8F5B61D-9152-4399-BF30-A1E4F3F072F6}#4.0#0"; "IGTabs40.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmDSSSetup 
   BorderStyle     =   0  'None
   Caption         =   "Larson Davis DSS Setup"
   ClientHeight    =   7200
   ClientLeft      =   240
   ClientTop       =   495
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdReload 
      Caption         =   "Reload"
      Height          =   555
      Left            =   480
      TabIndex        =   8
      Top             =   6480
      Width           =   1035
   End
   Begin VB.CommandButton cmdTestSampleRate 
      Caption         =   "Test Sample Rate"
      Height          =   495
      Left            =   5880
      TabIndex        =   7
      Top             =   6510
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdGetBoardInfo 
      Caption         =   "Get Board Information"
      Height          =   555
      Left            =   3720
      TabIndex        =   6
      Top             =   6480
      Width           =   1995
   End
   Begin ActiveTabs.SSActiveTabs sstabDSSSetup 
      Height          =   5835
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   10292
      _Version        =   262144
      TabCount        =   6
      TagVariant      =   ""
      Tabs            =   "frmDSSSetup.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   5445
         Left            =   30
         TabIndex        =   27
         Top             =   360
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   9604
         _Version        =   262144
         TabGuid         =   "frmDSSSetup.frx":0173
         Begin VB.CommandButton cmdGetStatus 
            Caption         =   "Get Status"
            Height          =   495
            Left            =   1560
            TabIndex        =   29
            Top             =   4200
            Width           =   1635
         End
         Begin VB.CommandButton cmdEnableAllChannels 
            Caption         =   "Enable All Channels"
            Height          =   495
            Left            =   3480
            TabIndex        =   28
            Top             =   4200
            Width           =   1755
         End
         Begin FPSpread.vaSpread vagrdStatus 
            Height          =   3855
            Left            =   60
            TabIndex        =   30
            Top             =   60
            Width           =   7275
            _Version        =   393216
            _ExtentX        =   12832
            _ExtentY        =   6800
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SpreadDesigner  =   "frmDSSSetup.frx":019B
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   5445
         Left            =   30
         TabIndex        =   25
         Top             =   360
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   9604
         _Version        =   262144
         TabGuid         =   "frmDSSSetup.frx":036F
         Begin FPSpread.vaSpread vagrdBoardInfo 
            Height          =   2655
            Left            =   60
            TabIndex        =   26
            Top             =   60
            Width           =   7695
            _Version        =   393216
            _ExtentX        =   13573
            _ExtentY        =   4683
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SpreadDesigner  =   "frmDSSSetup.frx":0397
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   5445
         Left            =   30
         TabIndex        =   9
         Top             =   360
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   9604
         _Version        =   262144
         TabGuid         =   "frmDSSSetup.frx":056B
         Begin VB.Frame fraPulseCounter 
            Caption         =   "Pulse Counter"
            Height          =   675
            Left            =   7140
            TabIndex        =   21
            Top             =   3060
            Width           =   2355
            Begin VB.CheckBox chkEnablePulse 
               Alignment       =   1  'Right Justify
               Caption         =   "Enable"
               Height          =   435
               Left            =   120
               TabIndex        =   22
               Top             =   180
               Width           =   2115
            End
         End
         Begin VB.Frame fraTachometer 
            Caption         =   "Tachometer"
            Height          =   1215
            Left            =   4620
            TabIndex        =   19
            Top             =   3060
            Width           =   2415
            Begin VB.TextBox txtPulsePerRev 
               Height          =   315
               Left            =   1500
               TabIndex        =   24
               Top             =   720
               Width           =   795
            End
            Begin VB.CheckBox chkEnableTach 
               Alignment       =   1  'Right Justify
               Caption         =   "Enable"
               Height          =   315
               Left            =   120
               TabIndex        =   20
               Top             =   300
               Width           =   1875
            End
            Begin VB.Label lblPulseRev 
               Caption         =   "Pulse/Rev:"
               Height          =   255
               Left            =   180
               TabIndex        =   23
               Top             =   780
               Width           =   1095
            End
         End
         Begin VB.Frame fraPowerSupply 
            Caption         =   "Power Supply"
            Height          =   615
            Left            =   2340
            TabIndex        =   17
            Top             =   3060
            Width           =   2235
            Begin VB.CheckBox chkPowerSupply 
               Alignment       =   1  'Right Justify
               Caption         =   "5.0 Volts"
               Height          =   255
               Left            =   180
               TabIndex        =   18
               Top             =   240
               Width           =   1875
            End
         End
         Begin VB.Frame fraIsolatedLogic 
            Caption         =   "Isolated Logic"
            Height          =   1755
            Left            =   180
            TabIndex        =   12
            Top             =   3060
            Width           =   2115
            Begin VB.CheckBox chkIsolatedLogicInput 
               Alignment       =   1  'Right Justify
               Caption         =   "Input A"
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   16
               Top             =   240
               Width           =   1755
            End
            Begin VB.CheckBox chkIsolatedLogicInput 
               Alignment       =   1  'Right Justify
               Caption         =   "Input B"
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   15
               Top             =   600
               Width           =   1755
            End
            Begin VB.CheckBox chkIsolatedLogicInput 
               Alignment       =   1  'Right Justify
               Caption         =   "Output Relay C"
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   14
               Top             =   960
               Width           =   1755
            End
            Begin VB.CheckBox chkIsolatedLogicInput 
               Alignment       =   1  'Right Justify
               Caption         =   "Output Relay D"
               Height          =   375
               Index           =   3
               Left            =   120
               TabIndex        =   13
               Top             =   1320
               Width           =   1755
            End
         End
         Begin FPSpread.vaSpread vagrdMaster 
            Height          =   2895
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   9555
            _Version        =   393216
            _ExtentX        =   16854
            _ExtentY        =   5106
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ScrollBars      =   2
            SpreadDesigner  =   "frmDSSSetup.frx":0593
         End
      End
      Begin ActiveTabs.SSActiveTabPanel sspnlAccel 
         Height          =   5445
         Left            =   30
         TabIndex        =   3
         Top             =   360
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   9604
         _Version        =   262144
         TabGuid         =   "frmDSSSetup.frx":0772
         Begin VB.CommandButton cmdLightLED 
            Caption         =   "Light LED"
            Height          =   435
            Left            =   300
            TabIndex        =   11
            Top             =   4860
            Width           =   1575
         End
         Begin FPSpread.vaSpread vagrdAccel 
            Height          =   4515
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   8955
            _Version        =   393216
            _ExtentX        =   15796
            _ExtentY        =   7964
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SpreadDesigner  =   "frmDSSSetup.frx":079A
         End
      End
      Begin ActiveTabs.SSActiveTabPanel sspnlTach 
         Height          =   5445
         Left            =   30
         TabIndex        =   2
         Top             =   360
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   9604
         _Version        =   262144
         TabGuid         =   "frmDSSSetup.frx":096E
      End
      Begin ActiveTabs.SSActiveTabPanel sspnlMic 
         Height          =   5445
         Left            =   30
         TabIndex        =   1
         Top             =   360
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   9604
         _Version        =   262144
         TabGuid         =   "frmDSSSetup.frx":0996
         Begin FPSpread.vaSpread vagrdMic 
            Height          =   5175
            Left            =   120
            TabIndex        =   4
            Top             =   120
            Width           =   9555
            _Version        =   393216
            _ExtentX        =   16854
            _ExtentY        =   9128
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SpreadDesigner  =   "frmDSSSetup.frx":09BE
         End
      End
   End
End
Attribute VB_Name = "frmDSSSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const m_cintBoardNumber As Integer = 1
Private Const m_cintBoardType As Integer = 2
Private Const m_cintRevision As Integer = 3
Private Const m_cintSerialNumber As Integer = 4

' Microphone Grid Column Locations
Private Const cintMicEnabled As Integer = 1
Private Const cintMicChannelNumber As Integer = 2
Private Const cintMicSampleRate As Integer = 3
Private Const cintMicBufferType As Integer = 4
Private Const cintMicRange As Integer = 5
Private Const cintMicModel As Integer = 6
Private Const cintMicSerialNum As Integer = 7
Private Const cintMicSensitivity As Integer = 8
Private Const cintMicDirection As Integer = 9
Private Const cintMicMicBias As Integer = 10
Private Const cintMicBoard As Integer = 11
Private Const cintMicChannel As Integer = 12

' Microphone Board Grid Column Locations
' Accel Grid Column Locations
Private Const cintAccelEnable As Integer = 1
Private Const cintAccelChannelNumber As Integer = 2
Private Const cintAccelSampleRate As Integer = 3
Private Const cintAccelBufferType As Integer = 4
Private Const cintAccelRange As Integer = 5
Private Const cintAccelModel As Integer = 6
Private Const cintAccelSerialNum As Integer = 7
Private Const cintAccelSensitivity As Integer = 8
Private Const cintAccelDirection As Integer = 9
Private Const cintAccelBoard As Integer = 10
Private Const cintAccelChannel As Integer = 11

' Master Grid Column Locations
Private Const cintMasterEnable As Integer = 1
Private Const cintMasterChannelNumber As Integer = 2
Private Const cintMasterChannelType As Integer = 3
Private Const cintMasterLabel As Integer = 4
Private Const cintMasterCalFactor As Integer = 5

' Max Columns
Private Const cintMicCols As Integer = 12
Private Const cintAccelCols As Integer = 11
Private Const cintMasterCols As Integer = 5
Private Const cintMasterRows As Integer = 10

' Board Information
Public Enum enuBoardInfo
    biModelNumber = 0
    biSerialNumber = 1
    biRevisionNumber = 2
    biLarsonDavis = 3
End Enum

' Channel Information
Public Enum enuChannelInfo
    ciChannelFunction = 0
    ciPhysicalChannelID = 1
    ciAdditionInfo = 2
End Enum

' Board Type Constants
Private Const BOARD_TYPE_MASTERECP = &H4D503031
Private Const BOARD_TYPE_MASTERNET = &H4D453031
Private Const BOARD_TYPE_DUALMIC = &H4D203031
Private Const BOARD_TYPE_DUALMICOPT1 = &H4D203032
Private Const BOARD_TYPE_QUADSOURCE = &H51533031
Private Const BOARD_TYPE_RECEIVER = &H52203031

' Channel Type Constants
Private Const CHANNEL_TYPE_GLOBAL = &H47423031
Private Const CHANNEL_TYPE_ADC = &H41443031
Private Const CHANNEL_TYPE_CNTR = &H43543031
Private Const CHANNEL_TYPE_OPTOIN = &H4F493031
Private Const CHANNEL_TYPE_OPTOOUT = &H4F4F3031
Private Const CHANNEL_TYPE_TACH = &H54413031
Private Const CHANNEL_TYPE_POWER = &H53563031
Private Const CHANNEL_TYPE_VOLTS = &H56543031
Private Const CHANNEL_TYPE_N1AUTO = &H414E3031
Private Const CHANNEL_TYPE_N1CROSS = &H434E3031
Private Const CHANNEL_TYPE_N3AUTO = &H414E3033
Private Const CHANNEL_TYPE_N3CROSS = &H434E3033
Private Const CHANNEL_TYPE_FAUTO = &H41463031
Private Const CHANNEL_TYPE_FCROSS = &H43463031
Private Const CHANNEL_TYPE_BOARDBAND_ACZ = &H42423031
Private Const CHANNEL_TYPE_BROADBAND_SUMMARY = &H42423032
Private Const CHANNEL_TYPE_DSIT = &H56543039
Private Const CHANNEL_TYPE_TACH_OR_TRIGGER = &H54413032
Private Const CHANNEL_TYPE_SIGGEN = &H53473031

' DSS Error Codes
Private Const DSS_EXERR_NONE = &H0
Private Const DSS_EXERR_CRC = &H1
Private Const DSS_EXERR_INVALID_COMMAND = &H2
Private Const DSS_EXERR_MODIFIED_COMMAND = &H3
Private Const DSS_EXERR_WRITE = &H1001
Private Const DSS_EXERR_READ = &H1002
Private Const DSS_EXERR_NOTCONNECTED = &H1003
Private Const DSS_EXERR_ALREADYCONNECTED = &H1004
Private Const DSS_EXERR_CANTCONNECT = &H1005
Private Const DSS_EXERR_UNKNOWNPROTOCOL = &H1006
Private Const DSS_EXERR_INVALIDCHANNEL = &H1007
Private Const DSS_EXERR_INVALIDPARAM = &H1008
Private Const DSS_EXERR_INVALIDCONTROL = &H1009
Private Const DSS_EXERR_PARAMOUTOFBOUND = &H100A
Private Const DSS_EXERR_FILE = &H100B
Private Const DSS_EXERR_MUSTSTOP = &H100C

' Master board control tags
Private Const TAG_TRIGGER_MODE As Long = &H54720000
Private Const TAG_TRIGGER_CONTROL As Long = &H54720003
Private Const TAG_TRIGGER_LEVEL As Long = &H54720001
Private Const TAG_TRIGGER_FREQUENCY As Long = &H54720002
Private Const TAG_PRETRIGGER_PERCENT  As Long = &H50720000
Private Const TAG_SAMPLE_RATE As Long = &H53610000
Private Const TAG_BUFFER_SIZE As Long = &H53610001
Private Const TAG_DISPLAY_BUFFER_SIZE As Long = &H44690000
Private Const TAG_SWITCHES As Long = &H53770000
Private Const TAG_COUNT_PERIOD As Long = &H436F0000
Private Const TAG_BUFFER_CONTROL As Long = &H53610002
    
' Signal generator board control tags
Private Const TAG_SIGNAL_GENERATOR_MODE As Long = &H47650001
Private Const TAG_SINE_MODE_FREQUENCIES As Long = &H53690003
Private Const TAG_SINE_MODE_LEVELS As Long = &H53690004
Private Const TAG_SINE_SWEEP_TIME_CONTROL As Long = &H53690005
Private Const TAG_BURST_MODE_CYCLE_COUNT As Long = &H53690002
Private Const TAG_NOISE_VOLTS As Long = &H4E6F0000
Private Const TAG_ARBITRARY_SCALE As Long = &H41720001
Private Const TAG_GATING As Long = &H47610001
Private Const TAG_PULSE_LEVELS As Long = &H50750001
Private Const TAG_PULSE_WIDTHS As Long = &H50750002
Private Const TAG_ARBWAVEFORM_FILENAME As Long = &H41720002
    
' Analog board control tags
Private Const TAG_GAIN = &H47610000
Private Const TAG_RANGE As Long = &H52610000
Private Const TAG_BIAS As Long = &H4D690000
Private Const TAG_PRE_EMPHASIS As Long = &H50720001
Private Const TAG_OCTAVE_BANDS As Long = &H4F630001
Private Const TAG_OCTAVE_LOW_FREQUENCY As Long = &H4F630002
Private Const TAG_THIRDOCTAVE_LOW_FREQUENCY As Long = &H4F630003
Private Const TAG_THIRDOCTAVE_BANDS As Long = &H4F630004
Private Const TAG_AVERAGING_MODE As Long = &H41760000
Private Const TAG_AVERAGING_SAMPLE_PERIOD As Long = &H41760001
Private Const TAG_FFT_LINES As Long = &H46460000
Private Const TAG_FFT_WINDOW As Long = &H46460001
Private Const TAG_FULL_CROSS_MODE As Long = &H4F630005
Private Const TAG_THIRD_CROSS_MODE As Long = &H4F630006
Private Const TAG_FFT_CROSS_MODE As Long = &H46460002
Private Const TAG_THIRD_ENVELOPE_ATTACK As Long = &H456E0001
Private Const TAG_THIRD_ENVELOPE_DECAY As Long = &H456E0002
Private Const TAG_THIRD_ENVELOPE_BAND As Long = &H456E0003
Private Const TAG_THIRD_ENVELOPE_INPUT As Long = &H456E0004
'private Const TAG_GATING As Long = &H47610001
Private Const TAG_OCTAVE_TRIGGER As Long = &H4F630007
Private Const TAG_THIRDOCTAVE_TRIGGER As Long = &H4F630008
Private Const TAG_FFT_TRIGGER As Long = &H46460003
Private Const TAG_BROADBAND_TRIGGER As Long = &H42420000
Private Const TAG_EXPONENTIAL_TIME_CONSTANT As Long = &H41760002

Private DSS As DSSSERVERLib.DSS
Private IMaster As DSSSERVERLib.IMaster
Private IDualMic As DSSSERVERLib.IDualMic
Private IReceiver As DSSSERVERLib.IReceiver

Private Sub Form_Load()
    
    '---- Temporary DSS Over-ride
    'ConnectToDSS
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'If IsConnected Then
    '    DSS.CloseConnection
    'End If
    
    Set DSS = Nothing
    Set IMaster = Nothing
    Set IDualMic = Nothing
End Sub


Private Sub ConnectToDSS()
    On Error GoTo ErrorHandler
    
    Dim lngLoaderMode As Long
    Dim varResults As Variant
    Dim blnConnected As Long
    
    
    AppLog InfoMsg, "ConnectToDSS,Connecting to DSS..."
    
    
    ' load the DSSServer and get interfaces
    Set DSS = New DSSSERVERLib.DSS
    Set IMaster = DSS
    Set IDualMic = DSS
    Set IReceiver = DSS

    ' Set the Default IP Address for the Connection Screen
    DSS.DefaultIPAddress = "192.168.1.99"
    DSS.DefaultIPAddressSelectMode = 1
    DSS.ConnectDlg True, False                                   'Connect to the DSS

    'InstrumentReady

    ' Is Connected
    DSS.IsConnected blnConnected

    ' Read the Board information
    DSS.InLoaderMode lngLoaderMode

    ' Check to see if in Loader Mode
    If lngLoaderMode = 1 Then
        ' Exit Loader Mode
        DSS.ExitLoaderMode
    End If
    
    'Get All the Board Information
    'GetAllBoardInfo

    InstrumentReady
    InitializeChannels

    ' Setup the Grids
    SetupMicGrid
    SetupAccelGrid
    SetupMaster
    GetBoards

    ' Get Microphone Settings
    GetMicrophoneSettings
    
    ' Get Accel Settings
    GetAccelSettings
    
    ' Get Master Settings
    GetMasterSettings

    With vagrdStatus
        .MaxCols = 10
        .MaxRows = 2
        .Col = 1
        .Row = 0
        .Text = "Board"
        .TypeHAlign = TypeHAlignCenter
        .Col = 2
        .Text = "Channel"
        .TypeHAlign = TypeHAlignCenter
        .Col = 3
        .Text = "Status"
        .TypeHAlign = TypeHAlignCenter
        .Col = 4
        .Text = "Available"
        .TypeHAlign = TypeHAlignCenter
        .Col = 5
        .Text = "Reachable"
        .TypeHAlign = TypeHAlignCenter
        .Col = 6
        .Text = "Enabled"
        .TypeHAlign = TypeHAlignCenter
        .Col = 7
        .Text = "# Control"
        .TypeHAlign = TypeHAlignCenter
        .Col = 8
        .Text = "Control Name"
        .TypeHAlign = TypeHAlignCenter
        .Col = 9
        .Text = "Control Type"
        .TypeHAlign = TypeHAlignCenter
        .Col = 10
        .Text = "Control Tag"
    End With
    
    Exit Sub
ErrorHandler:
    
End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub


Private Sub cmdGetBoardInfo_Click()

    ' Get Microphone Settings
    GetMicrophoneSettings
    
    ' Get Accel Settings
    GetAccelSettings
    
    ' Get Master Settings
    GetMasterSettings

End Sub

Private Sub cmdLightLED_Click()

    Dim intRow As Integer
    Dim bytBoard As Byte
    Dim bytChannel As Byte
    
    With vagrdAccel
        intRow = .SelBlockRow
        .Row = intRow
        .Col = cintAccelBoard
        bytBoard = CByte(.Value)
        .Col = cintAccelChannel
        bytChannel = CByte(.Value)
        
        DSS.EnableChannel bytBoard, bytChannel, 1
        DSS.EnableLED bytBoard, bytChannel
    End With
    

End Sub

Private Sub cmdOK_Click()
    
    ' Send Setting to the DSS
    SendMicSetting
    SendAccelSetting
    SendMasterSetting
    
    Unload Me

End Sub

Private Sub cmdReload_Click()

    DSS.Reconfigure
    DSS.InitChannelData
    DSS.UpdateTEDS

End Sub


Private Sub cmdTestSampleRate_Click()

    MicSetRange 1, 0, 2

End Sub

Private Sub SetupMicGrid()

    With vagrdMic
        ' Set the Number Columns
        .MaxCols = cintMicCols
        ' Set up Header Row Labels
        .Row = 0
        .Col = cintMicEnabled
        .Text = "Enabled"
        .TypeHAlign = TypeHAlignCenter
        .Col = cintMicChannelNumber
        .Text = "Channel #"
        .TypeHAlign = TypeHAlignCenter
        .Col = cintMicSampleRate
        .Text = "Sample Rate"
        .TypeHAlign = TypeHAlignCenter
        .Col = cintMicBufferType
        .Text = "Buffer Type"
        .TypeHAlign = TypeHAlignCenter
        .Col = cintMicRange
        .Text = "Range"
        .TypeHAlign = TypeHAlignCenter
        .Col = cintMicModel
        .Text = "Model"
        .TypeHAlign = TypeHAlignCenter
        .Col = cintMicSerialNum
        .Text = "Serial Number"
        .TypeHAlign = TypeHAlignCenter
        .Col = cintMicSensitivity
        .Text = "Sensitivity"
        .TypeHAlign = TypeHAlignCenter
        .Col = cintMicDirection
        .Text = "Direction"
        .TypeHAlign = TypeHAlignCenter
        .Col = cintMicMicBias
        .Text = "Mic Bias"
        .TypeHAlign = TypeHAlignCenter
    
        ' Resize the Column Width
        .RowHeight(0) = .MaxTextRowHeight(0)
        .ColWidth(cintMicEnabled) = .MaxTextColWidth(cintMicEnabled) + 2
        .ColWidth(cintMicChannelNumber) = .MaxTextColWidth(cintMicChannelNumber) + 2
        .ColWidth(cintMicSampleRate) = .MaxTextColWidth(cintMicSampleRate) + 2
        .ColWidth(cintMicBufferType) = .MaxTextColWidth(cintMicBufferType) + 2
        .ColWidth(cintMicRange) = .MaxTextColWidth(cintMicRange) + 2
        .ColWidth(cintMicModel) = .MaxTextColWidth(cintMicModel) + 2
        .ColWidth(cintMicSerialNum) = .MaxTextColWidth(cintMicSerialNum) + 2
        .ColWidth(cintMicSensitivity) = .MaxTextColWidth(cintMicSensitivity) + 2
        .ColWidth(cintMicDirection) = .MaxTextColWidth(cintMicDirection) + 2
        .ColWidth(cintMicMicBias) = .MaxTextColWidth(cintMicMicBias) + 2
        .ColWidth(cintMicBoard) = 0
        .ColWidth(cintMicChannel) = 0
    End With
            
End Sub

Private Sub SetupAccelGrid()

    With vagrdAccel
        ' Set the Number Columns
        .MaxCols = cintAccelCols
        ' Set up Header Row Labels
        .Row = 0
        .Col = cintAccelEnable
        .Text = "Enable"
        .TypeHAlign = TypeHAlignCenter
        .Col = cintAccelChannelNumber
        .Text = "Channel #"
        .TypeHAlign = TypeHAlignCenter
        .Col = cintAccelSampleRate
        .Text = "Sample Rate"
        .TypeHAlign = TypeHAlignCenter
        .Col = cintAccelBufferType
        .Text = "Buffer Type"
        .TypeHAlign = TypeHAlignCenter
        .Col = cintAccelRange
        .Text = "Range"
        .TypeHAlign = TypeHAlignCenter
        .Col = cintAccelModel
        .Text = "Model"
        .TypeHAlign = TypeHAlignCenter
        .Col = cintAccelSerialNum
        .Text = "Serial #"
        .TypeHAlign = TypeHAlignCenter
        .Col = cintAccelSensitivity
        .Text = "Sensitivity"
        .TypeHAlign = TypeHAlignCenter
        .Col = cintAccelDirection
        .Text = "Direction"
        .TypeHAlign = TypeHAlignCenter
    
        ' Resize the Column Width
        .RowHeight(0) = .MaxTextRowHeight(0)
        .ColWidth(cintAccelEnable) = .MaxTextColWidth(cintAccelEnable) + 2
        .ColWidth(cintAccelChannelNumber) = .MaxTextColWidth(cintAccelChannelNumber) + 2
        .ColWidth(cintAccelRange) = .MaxTextColWidth(cintAccelRange) + 2
        .ColWidth(cintAccelSampleRate) = .MaxTextColWidth(cintAccelSampleRate) + 2
        .ColWidth(cintAccelBufferType) = .MaxTextColWidth(cintAccelBufferType) + 2
        .ColWidth(cintAccelModel) = .MaxTextColWidth(cintAccelModel) + 2
        .ColWidth(cintAccelSerialNum) = .MaxTextColWidth(cintAccelSerialNum) + 2
        .ColWidth(cintAccelSensitivity) = .MaxTextColWidth(cintAccelSensitivity) + 2
        .ColWidth(cintAccelDirection) = .MaxTextColWidth(cintAccelDirection) + 2
        .ColWidth(cintAccelBoard) = 0
        .ColWidth(cintAccelChannel) = 0
    End With

End Sub

Private Sub Form_Resize()
    On Error Resume Next

    sstabDSSSetup.Width = Me.ScaleWidth - (sstabDSSSetup.Left * 2)
    
    vagrdMic.Width = sstabDSSSetup.Width - (vagrdMic.Left * 2)
    vagrdAccel.Width = sstabDSSSetup.Width - (vagrdAccel.Left * 2)

End Sub

Private Sub GetMicrophoneSettings()

    Dim bytBoardIndex As Byte
    Dim lngNumberBoards As Long
    Dim bytChannelIndex As Byte
    Dim lngNumberChannels As Long
    Dim lngRowIndex As Long
    Dim lngNumber As Long
    
    lngNumberBoards = GetNumberBoards
    lngRowIndex = 1
    
    For bytBoardIndex = 0 To lngNumberBoards - 1
        If Val(GetBoardType(bytBoardIndex, True)) = BOARD_TYPE_DUALMIC Then
            lngNumberChannels = GetNumberChannels(bytBoardIndex)
            For bytChannelIndex = 0 To lngNumberChannels - 1
                With vagrdMic
                    .MaxRows = lngRowIndex
                    .Row = lngRowIndex
                    ' Channel Enable
                    .Col = cintMicEnabled
                    .CellType = CellTypeCheckBox
                    If ChannelEnabled(bytBoardIndex, bytChannelIndex) Then
                        .Text = 1
                    Else
                        .Text = 0
                    End If
                    .TypeHAlign = TypeHAlignCenter
                    ' Channel
                    .Col = cintMicChannelNumber
'                    .Text = "Board " & bytBoardIndex & " Channel " & bytChannelIndex
                    .Text = GetChannelType(bytBoardIndex, bytChannelIndex)
                    ' Sample Rate
                    .Col = cintMicSampleRate
                    FillSampleRateCombo .Col, .Row, BOARD_TYPE_DUALMIC, vagrdMic
                    lngNumber = MicGetSampleRate(bytBoardIndex, bytChannelIndex)
                    If lngNumber = -1 Then
                        .Text = ""
                        .Lock = True
                        .BackColor = vbGrayText
                    Else
                        .Lock = False
                        .BackColor = vbWhite
                        .TypeComboBoxCurSel = lngNumber
                    End If
                    ' Buffer Type
                    .Col = cintMicBufferType
                    FillBufferTypeCombo .Col, .Row, vagrdMic
                    lngNumber = MicGetBuffertype(bytBoardIndex, bytChannelIndex)
                    If lngNumber = -1 Then
                        .Lock = True
                        .BackColor = vbGrayText
                    Else
                        .Lock = False
                        .BackColor = vbWhite
                        .TypeComboBoxCurSel = lngNumber
                    End If
                    ' Range
                    .Col = cintMicRange
                    FillMicRangeCombo .Col, .Row
                    lngNumber = MicGetRange(bytBoardIndex, bytChannelIndex)
                    If lngNumber = -1 Then
                        .Lock = True
                        .BackColor = vbGrayText
                    Else
                        .Lock = False
                        .BackColor = vbWhite
                        .TypeComboBoxCurSel = CInt(lngNumber)
                    End If
                    ' Mic Bias
                    .Col = cintMicMicBias
                    FillMicMicBiasCombo .Col, .Row
                    lngNumber = MicGetBias(bytBoardIndex, bytChannelIndex)
                    If lngNumber = -1 Then
                        .Lock = True
                        .BackColor = vbGrayText
                    Else
                        .Lock = False
                        .BackColor = vbWhite
                        .TypeComboBoxCurSel = CInt(lngNumber)
                    End If
                    ' Board Number
                    .Col = cintMicBoard
                    .Text = bytBoardIndex
                    ' Channel Number
                    .Col = cintMicChannel
                    .Text = bytChannelIndex
                End With
                lngRowIndex = lngRowIndex + 1
            Next bytChannelIndex
        End If
    Next bytBoardIndex

    ' Resize the Columns
    With vagrdMic
        .ColWidth(cintMicChannelNumber) = .MaxTextColWidth(cintMicChannelNumber) + 2
        .ColWidth(cintMicRange) = .MaxTextColWidth(cintMicRange) + 2
    End With

End Sub


Private Sub GetAccelSettings()

    Dim bytBoardIndex As Byte
    Dim lngNumberBoards As Long
    Dim bytChannelIndex As Byte
    Dim lngNumberChannels As Long
    Dim lngRowIndex As Long
    Dim lngNumber As Long
    
    lngNumberBoards = GetNumberBoards
    lngRowIndex = 1
    
    For bytBoardIndex = 0 To lngNumberBoards - 1
        If Val(GetBoardType(bytBoardIndex, True)) = BOARD_TYPE_RECEIVER Then
            lngNumberChannels = GetNumberChannels(bytBoardIndex)
            For bytChannelIndex = 0 To lngNumberChannels - 1
                With vagrdAccel
                    .MaxRows = lngRowIndex
                    .Row = lngRowIndex
                    ' Enable
                    .Col = cintAccelEnable
                    .CellType = CellTypeCheckBox
                    .Value = ChannelEnabled(bytBoardIndex, bytChannelIndex)
                    .TypeHAlign = TypeHAlignCenter
                    ' Channel
                    .Col = cintAccelChannelNumber
                    .Text = "Board " & bytBoardIndex & " Channel " & bytChannelIndex
                    ' Sample Rate
                    .Col = cintAccelSampleRate
                    FillSampleRateCombo .Col, .Row, BOARD_TYPE_RECEIVER, vagrdAccel
                    lngNumber = ReceiverGetSampleRate(bytBoardIndex, bytChannelIndex)
                    If lngNumber = -1 Then
                        .Text = ""
                        .Lock = True
                        .BackColor = vbGrayText
                    Else
                        .Lock = False
                        .BackColor = vbWhite
                        .TypeComboBoxCurSel = lngNumber
                    End If
                    ' Buffer Type
                    .Col = cintAccelBufferType
                    FillBufferTypeCombo .Col, .Row, vagrdAccel
                    lngNumber = ReceiverGetBuffertype(bytBoardIndex, bytChannelIndex)
                    If lngNumber = -1 Then
                        .Lock = True
                        .BackColor = vbGrayText
                    Else
                        .Lock = False
                        .BackColor = vbWhite
                        .TypeComboBoxCurSel = lngNumber
                    End If
                    ' Range
'                    .Col = cintAccelRange
'                    Call FillMicRangeCombo(.Col, .Row)
'                    lngNumber = ReceiverGetRange(bytBoardIndex, bytChannelIndex)
'                    If lngNumber = -1 Then
'                        .Lock = True
'                        .BackColor = vbGrayText
'                    Else
'                        .Lock = False
'                        .BackColor = vbWhite
'                        .TypeComboBoxCurSel = CInt(lngNumber)
'                    End If
                    .Col = cintAccelModel
                    '.Text = ReadControlTEDSSize(bytBoardIndex, bytChannelIndex)
                    ReadControlTEDS bytBoardIndex, bytChannelIndex
                    .Col = cintAccelBoard
                    .Text = bytBoardIndex
                    .Col = cintAccelChannel
                    .Text = bytChannelIndex
                End With
                lngRowIndex = lngRowIndex + 1
            Next bytChannelIndex
        End If
    Next bytBoardIndex

    With vagrdAccel
        .ColWidth(cintAccelRange) = .MaxTextColWidth(cintAccelRange) + 2
        .ColWidth(cintAccelChannelNumber) = .MaxTextColWidth(cintAccelChannelNumber) + 2
    End With

End Sub

Private Sub FillMicRangeCombo(ByVal intColumn As Integer, ByVal intRow As Integer)

    With vagrdMic
        .Row = intRow
        .Col = intColumn
        .CellType = CellTypeComboBox
        .TypeComboBoxClear intColumn, intRow
        .TypeComboBoxAutoSearch = TypeComboBoxAutoSearchNone
        .TypeComboBoxEditable = False
        .TypeComboBoxList = "14.2000 volts" & Chr$(9) + " 4.4900 volts" + Chr$(9) + _
                            " 1.4200 volts" & Chr$(9) + " 0.4490 volts" + Chr$(9) + _
                            " 0.1420 volts" & Chr$(9) + " 0.0449 volts"
    End With

End Sub

Private Sub FillSampleRateCombo(ByVal intColumn As Integer, _
                                ByVal intRow As Integer, _
                                ByVal lngBoard As Long, _
                                ByVal conGrid As Control)

    With conGrid
        .Row = intRow
        .Col = intColumn
        .CellType = CellTypeComboBox
        .TypeComboBoxClear intColumn, intRow
        .TypeComboBoxAutoSearch = TypeComboBoxAutoSearchNone
        .TypeComboBoxEditable = False
        Select Case lngBoard
            ' Master Boards
            Case BOARD_TYPE_MASTERECP, BOARD_TYPE_MASTERNET
                .TypeComboBoxList = "  100.0 Hz"
            ' Analog Boards
            Case BOARD_TYPE_DUALMIC, BOARD_TYPE_DUALMICOPT1
                .TypeComboBoxList = "61440.0 Hz"
            ' Source Board
            Case BOARD_TYPE_QUADSOURCE
                .TypeComboBoxList = "61440.0 Hz"
            ' Receiver Board
            Case BOARD_TYPE_RECEIVER
                ' DSIT
                    .TypeComboBoxList = " 5120.0 Hz" & Chr$(9) + _
                                        " 2560.0 Hz" + Chr$(9) + _
                                        " 1280.0 Hz" & Chr$(9) + _
                                        "  640.0 Hz" + Chr$(9) + _
                                        "  320.0 Hz" & Chr$(9) + _
                                        "  160.0 Hz"
                ' Tach
                    .TypeComboBoxList = "  100.0 Hz"
        End Select
    End With

End Sub

Private Sub FillBufferTypeCombo(ByVal intColumn As Integer, _
                                    ByVal intRow As Integer, _
                                    ByVal conGrid As Control)

    With conGrid
        .Row = intRow
        .Col = intColumn
        .CellType = CellTypeComboBox
        .TypeComboBoxClear intColumn, intRow
        .TypeComboBoxAutoSearch = TypeComboBoxAutoSearchNone
        .TypeComboBoxEditable = False
        .TypeComboBoxList = "Stop On Fill" & Chr$(9) & _
                            "Circular" + Chr$(9) + _
                            "Circular Pretrigger"
    End With

End Sub

Private Sub FillMicMicBiasCombo(ByVal intColumn As Integer, ByVal intRow As Integer)

    With vagrdMic
        .Row = intRow
        .Col = intColumn
        .CellType = CellTypeComboBox
        .TypeComboBoxClear intColumn, intRow
        .TypeComboBoxAutoSearch = TypeComboBoxAutoSearchNone
        .TypeComboBoxEditable = False
        .TypeComboBoxList = "  0.0 Volts" & Chr$(9) & _
                            " 28.0 Volts" + Chr$(9) + _
                            "200.0 Volts"
    End With

End Sub


Private Sub SetupMaster()

    Dim intRowIndex As Integer
    
    ' Setup Master A/D Channels
    With vagrdMaster
        .MaxRows = cintMasterRows
        .MaxCols = cintMasterCols
        .Row = 0
        ' Enable Column
        .Col = cintMasterEnable
        .TypeHAlign = TypeHAlignCenter
        .Text = "Enabled"
        ' Channel Number
        .Col = cintMasterChannelNumber
        .Text = "Channel #"
        .TypeHAlign = TypeHAlignCenter
        ' Channel Type
        .Col = cintMasterChannelType
        .Text = "Channel Type"
        .TypeHAlign = TypeHAlignCenter
        ' Channel Label
        .Col = cintMasterLabel
        .Text = "Channel Label"
        .TypeHAlign = TypeHAlignCenter
        ' Channel Cal Factor
        .Col = cintMasterCalFactor
        .Text = "Cal. Factor"
        .TypeHAlign = TypeHAlignCenter
        For intRowIndex = 1 To cintMasterRows
            .Row = intRowIndex
            ' Enabled Column
            .Col = cintMasterEnable
            .CellType = CellTypeCheckBox
            .TypeHAlign = TypeHAlignCenter
            ' Channel Number
            .Col = cintMasterChannelNumber
            .Text = intRowIndex
            .TypeHAlign = TypeHAlignCenter
            ' Channel Type
            .Col = cintMasterChannelType
            .Text = GetChannelInformation(0, CByte(intRowIndex), ciChannelFunction)
            .TypeHAlign = TypeHAlignCenter
            ' Channel Label
            .Col = cintMasterLabel
            ' Channel Cal Factor
            .Col = cintMasterCalFactor
        Next
        
        ' Resize the Column Width
        .RowHeight(0) = .MaxTextRowHeight(0)
        .ColWidth(cintMasterChannelType) = .MaxTextColWidth(cintMasterChannelType) + 2
        .ColWidth(cintMasterLabel) = .MaxTextColWidth(cintMasterLabel) + 2
        .ColWidth(cintMasterCalFactor) = .MaxTextColWidth(cintMasterCalFactor) + 2
    End With

End Sub

Private Sub GetMasterSettings()

    Dim bytBoardIndex As Byte
    Dim lngNumberBoards As Long
    Dim bytChannelIndex As Byte
    Dim lngNumberChannels As Long
    Dim lngRowIndex As Long
    Dim lngNumber As Long
    
    lngNumberBoards = GetNumberBoards
    lngRowIndex = 1
    
    ' Find the Master Board
    For bytBoardIndex = 0 To lngNumberBoards - 1
        If Val(GetBoardType(bytBoardIndex, True)) = BOARD_TYPE_MASTERNET Then
            ' Found Master Net Board
            Exit For
        End If
    Next bytBoardIndex
    
    ' Get the number channels on the board
    lngNumberChannels = GetNumberChannels(bytBoardIndex)
    
    ' 0 to 5 volt Channels
    For bytChannelIndex = 1 To 10
        With vagrdMaster
            .Row = bytChannelIndex
            ' Channel Enabled/Disabled
            .Col = cintMasterEnable
            .Value = ChannelEnabled(bytBoardIndex, bytChannelIndex)
        End With
    Next bytChannelIndex
    
    ' Power supply
    If ChannelEnabled(bytBoardIndex, 11) Then
        chkPowerSupply.Value = vbChecked
    Else
        chkPowerSupply.Value = vbUnchecked
    End If
    
    ' Isolated Logic Inputs
    For bytChannelIndex = 12 To 13
        If ChannelEnabled(bytBoardIndex, bytChannelIndex) Then
            chkIsolatedLogicInput(bytChannelIndex - 12).Value = vbChecked
        Else
            chkIsolatedLogicInput(bytChannelIndex - 12).Value = vbUnchecked
        End If
    Next
    
    ' Tach
    If ChannelEnabled(bytBoardIndex, 14) Then
        chkEnableTach.Value = vbChecked
    Else
        chkEnableTach.Value = vbUnchecked
    End If
    
    ' Pulse Counter
    If ChannelEnabled(bytBoardIndex, 15) Then
        chkEnablePulse.Value = vbChecked
    Else
        chkEnablePulse.Value = vbUnchecked
    End If

    ' Isolated Logic Outputs
    For bytChannelIndex = 16 To 17
        If ChannelEnabled(bytBoardIndex, bytChannelIndex) Then
            chkIsolatedLogicInput(bytChannelIndex - 14).Value = vbChecked
        Else
            chkIsolatedLogicInput(bytChannelIndex - 14).Value = vbUnchecked
        End If
    Next

End Sub



Private Sub SendMicSetting()
    
    Dim intRowIndex As Integer
    Dim bytBoard As Byte
    Dim bytChannel As Byte

    ' Microphone Setting
    With vagrdMic
        For intRowIndex = 1 To .MaxRows
            ' Set the Row
            .Row = intRowIndex
            ' Get the Board and Channel information
            .Col = cintMicBoard
            bytBoard = CByte(.Text)
            .Col = cintMicChannel
            bytChannel = CByte(.Text)
            ' Enable/Disable Channel
            .Col = cintMicEnabled
            If .Value = 1 Then
                DSS.EnableChannel bytBoard, bytChannel, 1
            Else
                DSS.EnableChannel bytBoard, bytChannel, 0
            End If
            ' Range
            .Col = cintMicRange
            If .Value <> "" Then
                MicSetRange bytBoard, bytChannel, .Value
            End If
            ' Buffer
            .Col = cintMicBufferType
            If .Value <> "" Then
                MicSetBufferType bytBoard, bytChannel, .Value
            End If
        Next
    End With

End Sub

Private Sub SendAccelSetting()
    
    Dim intRowIndex As Integer
    Dim bytBoard As Byte
    Dim bytChannel As Byte

    ' Accelerometer Setting
    With vagrdAccel
        For intRowIndex = 1 To .MaxRows
            ' Set the Row
            .Row = intRowIndex
            ' Get the Board and Channel information
            .Col = cintAccelBoard
            bytBoard = CByte(.Text)
            .Col = cintAccelChannel
            bytChannel = CByte(.Text)
            ' Enable/Disable Channel
            .Col = cintAccelEnable
            If .Value = 1 Then
                DSS.EnableChannel bytBoard, bytChannel, 1
            Else
                DSS.EnableChannel bytBoard, bytChannel, 0
            End If
            ' Range
            .Col = cintAccelRange
            If .Value <> "" Then
                MicSetRange bytBoard, bytChannel, .Value
            End If
            ' Buffer
            .Col = cintAccelBufferType
            If .Value <> "" Then
                MicSetBufferType bytBoard, bytChannel, .Value
            End If
        Next
    End With

End Sub

Private Sub SendMasterSetting()

    Dim intRowIndex As Integer
    Dim bytBoard As Byte
    Dim bytChannel As Byte
    Dim bytChannelIndex As Byte

    ' Master Setting
    With vagrdMaster
        For intRowIndex = 1 To .MaxRows
            ' Set the Row
            .Row = intRowIndex
            ' Master Board is assumed to be first
            bytBoard = 0
            ' Enable Channel
            .Col = cintMasterEnable
            If .Value = 1 Then
                DSS.EnableChannel bytBoard, CByte(intRowIndex), 1
            Else
                DSS.EnableChannel bytBoard, CByte(intRowIndex), 0
            End If
        Next
    End With
    
    ' Power supply
    If chkPowerSupply.Value = vbChecked Then
        DSS.EnableChannel bytBoard, 11, 1
    Else
        DSS.EnableChannel bytBoard, 11, 0
    End If
    
    ' Isolated Logic Inputs
    For bytChannelIndex = 12 To 13
        If chkIsolatedLogicInput(bytChannelIndex - 12).Value = vbChecked Then
            DSS.EnableChannel bytBoard, bytChannelIndex, 1
        Else
            DSS.EnableChannel bytBoard, bytChannelIndex, 0
        End If
    Next
    
    ' Tach
    If chkEnableTach.Value = vbChecked Then
        DSS.EnableChannel bytBoard, 14, 1
    Else
        DSS.EnableChannel bytBoard, 14, 0
    End If
    
    ' Pulse Counter
    If chkEnablePulse.Value = vbChecked Then
        DSS.EnableChannel bytBoard, 15, 1
    Else
        DSS.EnableChannel bytBoard, 15, 0
    End If

    ' Isolated Logic Outputs
    For bytChannelIndex = 16 To 17
        If chkIsolatedLogicInput(bytChannelIndex - 14).Value = vbChecked Then
            DSS.EnableChannel bytBoard, bytChannelIndex, 1
        Else
            DSS.EnableChannel bytBoard, bytChannelIndex, 0
        End If
    Next
    
    
End Sub



Private Sub GetBoards()
        On Error GoTo Form_Load_Err

        Dim intIndex As Integer
        Dim lngWidth As Long
        Dim lngTotalWidth As Long
    
        ' Format Grid
100     FormatGrid
        ' Fill in the Grid with Module Information
102     GetBoardInfo
    
104     With vagrdBoardInfo
            ' Grid Width
106         lngTotalWidth = 0
108         For intIndex = 0 To 4
110             .ColWidthToTwips .ColWidth(intIndex), lngWidth
112             lngTotalWidth = lngTotalWidth + lngWidth
            Next
114         .Width = lngTotalWidth
            ' Grid Height
116         lngTotalWidth = 0
118         For intIndex = 0 To 4
120             .RowHeightToTwips intIndex, .RowHeight(intIndex), lngWidth
122             lngTotalWidth = lngTotalWidth + lngWidth
            Next
124         .Height = lngTotalWidth
    
            ' Form Width
126         .Left = 100
128         Me.Width = .Width + (.Left * 2) + 120
            ' Form Height
        End With
        
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in MTUProject.frmDSSBoards.Form_Load " & _
               "at line " & Erl
        Resume Next
End Sub

Private Sub FormatGrid()
        On Error GoTo FormatGrid_Err

100     With vagrdBoardInfo
            ' Remove Scroll Bars
102         .ScrollBars = ScrollBarsNone
        
104         .MaxCols = 4
106         .MaxRows = GetNumberBoards
            ' Header Text
108         .Row = 0
            ' Board Number
110         .Col = m_cintBoardNumber
112         .Text = "Board #"
114         .TypeHAlign = TypeHAlignCenter
            ' Board Type
116         .Col = m_cintBoardType
118         .Text = "Board Type"
120         .TypeHAlign = TypeHAlignCenter
            ' Revision
122         .Col = m_cintRevision
124         .Text = "Revision"
126         .TypeHAlign = TypeHAlignCenter
            ' Serial #
128         .Col = m_cintSerialNumber
130         .Text = "Serial #"
132         .TypeHAlign = TypeHAlignCenter
    
            ' Resize Header
134         .RowHeight(0) = .MaxTextRowHeight(0)
        End With

        Exit Sub

FormatGrid_Err:
        MsgBox Err.Description & vbCrLf & _
               "in MTUProject.frmDSSBoards.FormatGrid " & _
               "at line " & Erl
        Resume Next
End Sub

Private Sub GetBoardInfo()
        On Error GoTo GetBoardInfo_Err

        Dim bytIndex As Byte
        Dim intNumberBoards As Integer
    
100     intNumberBoards = GetNumberBoards - 1
    
102     With vagrdBoardInfo
104         For bytIndex = 0 To intNumberBoards
106             .Row = bytIndex + 1
                ' Board Number
108             .Col = m_cintBoardNumber
110             .Text = bytIndex
112             .TypeHAlign = TypeHAlignCenter
                ' Board Type
114             .Col = m_cintBoardType
116             .Text = GetBoardType(bytIndex)
                ' Board Revision
118             .Col = m_cintRevision
120             .Text = GetBoardInformation(bytIndex, biRevisionNumber)
122             .TypeHAlign = TypeHAlignCenter
                ' Board Serial #
124             .Col = m_cintSerialNumber
126             .Text = GetBoardInformation(bytIndex, biSerialNumber)
128             .TypeHAlign = TypeHAlignCenter
            Next
        
            ' Resize the Columns
130         .ColWidth(m_cintBoardNumber) = .MaxTextColWidth(m_cintBoardNumber) + 2
132         .ColWidth(m_cintBoardType) = .MaxTextColWidth(m_cintBoardType) + 2
134         .ColWidth(m_cintRevision) = .MaxTextColWidth(m_cintRevision) + 2
136         .ColWidth(m_cintSerialNumber) = .MaxTextColWidth(m_cintSerialNumber) + 2
        End With

        Exit Sub
GetBoardInfo_Err:
        MsgBox Err.Description & vbCrLf & _
               "in MTUProject.frmDSSBoards.GetBoardInfo " & _
               "at line " & Erl
        Resume Next
End Sub


Private Sub cmdEnableAllChannels_Click()

    Dim bytBoardIndex As Byte
    Dim bytChannelIndex As Byte
    Dim bytNumberBoards As Byte
    Dim bytNumberChannels As Byte
    
    bytNumberBoards = GetNumberBoards
    
    For bytBoardIndex = 0 To bytNumberBoards - 1
        bytNumberChannels = GetNumberChannels(bytBoardIndex)
        For bytChannelIndex = 0 To bytNumberChannels - 1
            DSS.EnableChannel bytBoardIndex, bytChannelIndex, 1
        Next bytChannelIndex
    Next bytBoardIndex

End Sub

Private Sub cmdGetStatus_Click()

    Dim lngNumberBoards As Long
    Dim lngNumberChannels As Long
    Dim bytBoardIndex As Byte
    Dim bytChannelIndex As Byte
    Dim intRowIndex As Integer
    Dim strBoardType As String
    Dim lngIndex As Long
    Dim lngNumber As Long
    
    ' Get number of Boards
    DSS.GetNumBoards lngNumberBoards
    
    intRowIndex = 1
    
    With vagrdStatus
        For bytBoardIndex = 0 To lngNumberBoards - 1
            ' Get the number of Channels for the board
            lngNumberChannels = GetNumberChannels(bytBoardIndex)
            ' Get the Board Name
            strBoardType = GetBoardType(bytBoardIndex)
            For bytChannelIndex = 0 To lngNumberChannels - 1
                ' Add Row
                .MaxRows = intRowIndex
                ' Set Cursor Location
                .Row = intRowIndex
                ' Board
                .Col = 1
                .Text = strBoardType
                ' Channel
                .Col = 2
                .Text = "Channel " & bytChannelIndex
                ' Status
                .Col = 3
                .Text = GetChannelStatusString(bytBoardIndex, bytChannelIndex)
                ' Available
                .Col = 4
                .CellType = CellTypeCheckBox
                .TypeHAlign = TypeHAlignCenter
                .Text = ChannelAvailable(bytBoardIndex, bytChannelIndex)
                ' Reachable
                .Col = 5
                .CellType = CellTypeCheckBox
                .TypeHAlign = TypeHAlignCenter
                .Text = ChannelReachable(bytBoardIndex, bytChannelIndex)
                ' Enabled
                .Col = 6
                .CellType = CellTypeCheckBox
                .TypeHAlign = TypeHAlignCenter
                .Text = ChannelEnabled(bytBoardIndex, bytChannelIndex)
                ' Control Count
                .Col = 7
                .TypeHAlign = TypeHAlignCenter
                .Text = GetChannelControlCount(bytBoardIndex, bytChannelIndex)
                ' Control Count
                lngNumber = Val(.Text)
                If lngNumber >= 1 Then
                    For lngIndex = 0 To lngNumber - 1
                        .Col = 8
                        If Len(.Text) > 0 Then
                            .Text = .Text & "; "
                        End If
                        .Text = .Text & GetChannelControlName(bytBoardIndex, bytChannelIndex, lngIndex)
                    Next lngIndex
                End If
                ' Control type
                If lngNumber >= 1 Then
                    For lngIndex = 0 To lngNumber - 1
                        .Col = 9
                        If Len(.Text) > 0 Then
                            .Text = .Text & "; "
                        End If
'                        .Text = .Text & GetChannelControlType(bytBoardIndex, bytChannelIndex, lngIndex)
                    Next lngIndex
                End If
                ' Control Tag
                If lngNumber >= 1 Then
                    For lngIndex = 0 To lngNumber - 1
                        .Col = 10
                        If Len(.Text) > 0 Then
                            .Text = .Text & "; "
                        End If
                        .Text = .Text & GetChannelControlTag(bytBoardIndex, bytChannelIndex, lngIndex)
                    Next lngIndex
                End If
                
                ' Increment Row Index
                intRowIndex = intRowIndex + 1
            Next bytChannelIndex
        Next bytBoardIndex
        ' Adjust Column Width
        .ColWidth(1) = .MaxTextColWidth(1) + 2
        .ColWidth(2) = .MaxTextColWidth(2) + 2
        .ColWidth(3) = .MaxTextColWidth(3) + 2
        .ColWidth(8) = .MaxTextColWidth(8) + 2
        .ColWidth(9) = .MaxTextColWidth(9) + 2
        .ColWidth(10) = .MaxTextColWidth(10) + 2
    End With
    
End Sub


Public Sub GetAllBoardInfo()

    Dim bytIndex As Byte
    Dim varResults As Variant
    Dim bytResults(0 To 1029) As Byte
    Dim strResults As String
    Dim lngErrorNumber As Long
    
    On Error GoTo ErrHandler
    
    varResults = bytResults
    
    For bytIndex = 0 To 3
        DSS.ReadBoardInfo bytIndex, varResults
    Next bytIndex
    
    Exit Sub
    
ErrHandler:

    MsgBox Err.Description
'    DSS.GetExtendedError lngErrorNumber
    MsgBox GetDSSError
    
End Sub

Public Function GetDSSError() As String

    Dim lngErrorNumber As Long
    Dim strErrorMessage As String

    DSS.GetExtendedError lngErrorNumber
    ' Change to Database Later
    Select Case lngErrorNumber
        ' None
        Case DSS_EXERR_NONE
            strErrorMessage = ""
        ' CRC Error
        Case DSS_EXERR_CRC
            strErrorMessage = "A CRC error occurred while sending or receiving a packet."
        Case DSS_EXERR_INVALID_COMMAND
            strErrorMessage = "An unrecognized command was sent to the DSS."
        Case DSS_EXERR_MODIFIED_COMMAND
            strErrorMessage = "A command was sent to the DSS but contained illegal data."
        Case DSS_EXERR_WRITE
            strErrorMessage = "An error occured while writing data to the DSS."
        Case DSS_EXERR_READ
            strErrorMessage = "An error occured while reading data from the DSS."
        Case DSS_EXERR_NOTCONNECTED
            strErrorMessage = "An attempt was made to communicate with a DSS that isn't connected."
        Case DSS_EXERR_ALREADYCONNECTED
            strErrorMessage = "An attempt was made to connect to a DSS that was already connected."
        Case DSS_EXERR_CANTCONNECT
            strErrorMessage = "An error occurred while attempting to connect to the DSS."
        Case DSS_EXERR_UNKNOWNPROTOCOL
            strErrorMessage = "An error occurred while attempteng to connect to the DSS using an invalid protocol."
        Case DSS_EXERR_INVALIDCHANNEL
            strErrorMessage = "An error occurred while reading or writing to an invalid channel."
        Case DSS_EXERR_INVALIDPARAM
            strErrorMessage = "An error was made to read or write data using an invalid buffer or parameter."
        Case DSS_EXERR_INVALIDCONTROL
            strErrorMessage = "An error occurred while reading or writing to an invalid control."
        Case DSS_EXERR_PARAMOUTOFBOUND
            strErrorMessage = "An error occurred while writing an invalid parameter to the DSS."
        Case DSS_EXERR_FILE
            strErrorMessage = "An error occurred while attempting to open or read a waveform file."
        Case DSS_EXERR_MUSTSTOP
            strErrorMessage = "An error occurred because the instrument was not in the stopped state."
        Case Else
            strErrorMessage = "Unknown Error - Error Number " & lngErrorNumber
    End Select

    GetDSSError = strErrorMessage

End Function

Public Function GetBoardFirmwareRev(bytBoard As Byte) As String

    Dim strRevNumber As String

    On Error GoTo ErrHandler
    
    strRevNumber = Space(8)
    
    DSS.ReadBoardFirmwareRev bytBoard, strRevNumber
    
    GetBoardFirmwareRev = strRevNumber

    Exit Function
    
ErrHandler:

    MsgBox Err.Description
'    DSS.GetExtendedError lngErrorNumber
    MsgBox GetDSSError

End Function

Public Function GetBoardType(bytBoard As Byte, _
                            Optional blnNumber As Boolean = False) As String

    Dim lngBoardType As Long

    On Error GoTo ErrHandler
        
    DSS.GetBoardType bytBoard, lngBoardType

    If Not blnNumber Then
        Select Case lngBoardType
            Case BOARD_TYPE_MASTERECP
                GetBoardType = "DDS Master Module Parallel-Port"
            Case BOARD_TYPE_MASTERNET
                GetBoardType = "DDS Master Module for Ethernet"
            Case BOARD_TYPE_DUALMIC
                GetBoardType = "DDS Dual Microphone Module"
            Case BOARD_TYPE_DUALMICOPT1
                GetBoardType = "DDS Dual Microphone Option 1 Module"
            Case BOARD_TYPE_QUADSOURCE
                GetBoardType = "DDS Quad Source"
            Case BOARD_TYPE_RECEIVER
                GetBoardType = "DDS Digital Receiver Board"
            Case Else
                GetBoardType = "Unknown Module"
        End Select
    Else
        GetBoardType = lngBoardType
    End If
    
    Exit Function
    
ErrHandler:

    MsgBox Err.Description
'    DSS.GetExtendedError lngErrorNumber
    MsgBox GetDSSError

End Function

Public Function GetBoardInformation(bytBoard As Byte, _
                            enuInformationType As enuBoardInfo) As String

    Dim strResults As String
    
    DSS.GetBoardString bytBoard, enuInformationType, strResults
    
    GetBoardInformation = strResults

End Function

Public Function GetChannelInformation(bytBoard As Byte, bytChannel As Byte, _
                            enuInformationType As enuChannelInfo) As String

    Dim strResults As String
    
    DSS.GetChannelString bytBoard, bytChannel, enuInformationType, strResults
    
    GetChannelInformation = strResults

End Function

Public Function GetChannelType(bytBoard As Byte, bytChannel As Byte) _
                            As String

    Dim lngChannelType As Long
    
    DSS.GetChannelType bytBoard, bytChannel, lngChannelType
    
    Select Case lngChannelType
        Case CHANNEL_TYPE_GLOBAL
            GetChannelType = "Global Controls"
        Case CHANNEL_TYPE_ADC
            GetChannelType = "Auxiliary Input 0 to 5 Volts"
        Case CHANNEL_TYPE_CNTR
            GetChannelType = "Pulse Counter"
        Case CHANNEL_TYPE_OPTOIN
            GetChannelType = "Isolated Logic Inputs"
        Case CHANNEL_TYPE_OPTOOUT
            GetChannelType = "Isolated Output Relay"
        Case CHANNEL_TYPE_TACH
            GetChannelType = "Tachometer Hz"
        Case CHANNEL_TYPE_POWER
            GetChannelType = "Power Supply Volts"
        Case CHANNEL_TYPE_VOLTS
            GetChannelType = "Voltage vs. Time"
        Case CHANNEL_TYPE_N1AUTO
            GetChannelType = "1/1 Auto Spectrum"
        Case CHANNEL_TYPE_N1CROSS
            GetChannelType = "1/1 Cross Spectrum"
        Case CHANNEL_TYPE_N3AUTO
            GetChannelType = "1/3 Auto Spectrum"
        Case CHANNEL_TYPE_N3CROSS
            GetChannelType = "1/3 Cross Spectrum"
        Case CHANNEL_TYPE_FAUTO
            GetChannelType = "FFT Auto Spectrum"
        Case CHANNEL_TYPE_FCROSS
            GetChannelType = "FFT Cross Spectrum"
        Case CHANNEL_TYPE_BOARDBAND_ACZ
            GetChannelType = "Broadband ACZ Channel"
        Case CHANNEL_TYPE_BROADBAND_SUMMARY
            GetChannelType = "Mic Summary Data"
        Case CHANNEL_TYPE_DSIT
            GetChannelType = "Digital Receiver Channel"
        Case CHANNEL_TYPE_TACH_OR_TRIGGER
            GetChannelType = "Tach or Trigger Channel"
        Case CHANNEL_TYPE_SIGGEN
            GetChannelType = "Signal Generator"
        Case Else
            GetChannelType = "Unknown Channel Type"
    End Select

End Function

Public Function GetNumberChannels(bytBoard As Byte) As Long

    Dim lngNumChannels As Long
    
    DSS.GetNumChannels bytBoard, lngNumChannels
    
    GetNumberChannels = lngNumChannels
    
End Function

Public Function GetChannelStatusString(bytBoard As Byte, bytChannel As Byte) As String

    Dim lngStatus As Long
    Dim strResults As String
    Dim bytResults() As Byte
    Dim intIndex As Integer
    
    DSS.ReadChannelStatus bytBoard, bytChannel, lngStatus
    
    lngStatus = lngStatus
    
    bytResults() = DecimalToBinary(lngStatus)
    
    strResults = ""
    
    ' Grab all the Status Messages
    For intIndex = 0 To 31
        If bytResults(intIndex) = 1 Then
            ' Add next message to the End and separate with a Comma
            If Len(strResults) > 0 Then
                strResults = strResults & "; "
            End If
            Select Case intIndex
                Case 0
                    strResults = strResults & "Service Request"
                Case 1
                    strResults = strResults & "Trigger Acknowledge"
                Case 2
                    strResults = strResults & "Reset"
                Case 4
                    strResults = strResults & "Set to indicate auxiliary status bits exist"
                Case 5
                    strResults = strResults & "Data Buffering Error"
                Case 6
                    strResults = strResults & "Data Buffer Full"
                Case 7
                    strResults = strResults & "Hardware Error"
                Case 8
                    strResults = strResults & "Run/Enable"
                Case 12
                    strResults = strResults & "DSIT, Tranducer or Board present"
                Case 13
                    strResults = strResults & "Channel Unreachable"
                Case 14
                    strResults = strResults & "Sensor TEDS Available"
                Case 15
                    strResults = strResults & "Sensor TEDS read error"
                Case 16
                    strResults = strResults & "IO Busy/Channel Busy"
                Case 17
                    strResults = strResults & "Calibration error"
                Case 18
                    strResults = strResults & "Self-Test fail"
                Case 19
                    strResults = strResults & "TestMode"
                Case 20
                    strResults = strResults & "Over Range"
                Case 21
                    strResults = strResults & "ICP Shorted"
                Case 22
                    strResults = strResults & "ICP Open"
                Case 3, 9, 10, 11, 23, 24, 25
                    strResults = strResults & "Reserved " & lngStatus
                Case 26, 27, 28, 29, 30, 31
                    strResults = strResults & "Open " & lngStatus
                Case Else
                    strResults = strResults & "Not Defined"
            End Select
        End If
    Next
    
    GetChannelStatusString = strResults
    
End Function

Public Function ChannelAvailable(bytBoard As Byte, bytChannel As Byte) As Boolean

    Dim blnResults As Long
    
    DSS.ChannelAvailable bytBoard, bytChannel, blnResults
    
    ChannelAvailable = blnResults

End Function

Public Sub InitializeChannels()
        On Error GoTo InitializeChannels_Err
        DSS.InitChannelData
        Exit Sub
InitializeChannels_Err:
        Err.Clear
        GetDSSError
        Resume Next
End Sub

Public Sub InstrumentReady()
        On Error GoTo InstrumentReady_Err
        Dim lngResults As Long
        Screen.MousePointer = vbHourglass
        DSS.InstrumentBusy lngResults
        If lngResults Then
            Do
                DSS.WaitForInstrumentNotBusy
                DSS.InstrumentBusy lngResults
            Loop Until Not lngResults
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
InstrumentReady_Err:
        MsgBox Err.Description & vbCrLf & _
               "in MTUProject.modDSS.InstrumentReady " & _
               "at line " & Erl
        Resume Next
End Sub

Public Function ChannelReachable(bytBoard As Byte, bytChannel As Byte) As Boolean
    Dim lngResults As Long
    DSS.ChannelReachable bytBoard, bytChannel, lngResults
    ChannelReachable = lngResults
End Function

Public Function ChannelEnabled(bytBoard As Byte, bytChannel As Byte) As Boolean
    Dim lngResults As Long
    DSS.ChannelEnabled bytBoard, bytChannel, lngResults
    ChannelEnabled = CBool(lngResults)
End Function

Public Function GetChannelControlCount(bytBoard As Byte, bytChannel As Byte) As Long
    Dim lngResults As Long
    DSS.ChannelEnabled bytBoard, bytChannel, lngResults
    GetChannelControlCount = lngResults
End Function

Public Function GetChannelControlName(bytBoard As Byte, bytChannel As Byte, lngIndex As Long) As String
    Dim strResults As String
    DSS.GetChannelControlName bytBoard, bytChannel, lngIndex, strResults
    GetChannelControlName = strResults
End Function

Public Function GetNumberBoards() As Byte
    Dim lngResults As Long
    DSS.GetNumBoards lngResults
    GetNumberBoards = lngResults
End Function

Public Function IsConnected() As Boolean
    Dim lngResults As Long
    DSS.IsConnected lngResults
    IsConnected = lngResults
End Function

Public Function GetChannelControlType(bytBoard As Byte, bytChannel As Byte, lngIndex As Long)
        On Error GoTo GetChannelControlType_Err

        Dim lngResults As Long
        
        lngResults = -1
    
100     DSS.GetChannelControlType bytBoard, bytChannel, lngIndex, lngResults
        
102     GetChannelControlType = lngResults

        Exit Function

GetChannelControlType_Err:
        MsgBox Err.Description & vbCrLf & _
               "in Project1.modDSS.GetChannelControlType " & _
               "at line " & Erl & vbCrLf & _
               GetDSSError
        Resume Next
End Function

Public Function ChannelSupportsControl(ByVal bytBoard As Byte, ByVal bytChannel As Byte, ByVal lngTag As Long) As Boolean
    Dim lngResults As Long
    DSS.ChannelSupportsControl bytBoard, bytChannel, lngTag, lngResults
    If lngResults = 1 Then
        ChannelSupportsControl = True
    Else
        ChannelSupportsControl = False
    End If
End Function

Public Function GetControlIndexFromTag(bytBoard As Byte, bytChannel As Byte, lngTag As Long) As Long
    Dim lngResults As Long
    DSS.GetControlIndexFromTag bytBoard, bytChannel, lngTag, lngResults
    GetControlIndexFromTag = lngResults
End Function

Public Function GetChannelControlTag(bytBoard As Byte, bytChannel As Byte, lngIndex As Long) As Long
    Dim lngResults As Long
    DSS.GetChannelControlTag bytBoard, bytChannel, lngIndex, lngResults
    GetChannelControlTag = lngResults
End Function

Public Sub Test()

    Dim board As Byte
    Dim channel As Byte
    Dim bufsize As Long
    Dim v As Variant
    Dim buf(0 To 255) As Double
    Dim q As String * 1024
    Dim lngLoaderMode As Long
    Dim varResults As Variant
    Dim lngNumChannels As Long
    Dim iValue As Long
        
    board = 2
    channel = 2
    bufsize = 256
    '    dispbufsize = 256
        
'    board = 0
'    channel = 1
    ' configure the channel
    DSS.EnableChannel board, channel, 1
'    IReceiver.SetBufferSize board, channel, bufsize
'    IReceiver.SetBufferControl board, channel, 0
        
    DSS.GetChannelString board, channel, 0, q
        
    DSS.Run
    ' wait for the instrument to finish
    DSS.Stop
        
    ' read the data from the channel
    v = buf
    DSS.ReadChannelData board, channel, v
            
    'With frmPlot1.ChartFX1(0)
'        .ClearData CD_DATA
    '    .Axis(AXIS_Y).AutoScale = True
    '    .OpenDataEx COD_VALUES Or COD_ADDPOINTS, 1, 256
    '    For iValue = 1 To .MaxValues
    '        ' Last value
    '        .Value(-1) = v(iValue)
    '    Next
    '    .CloseData COD_VALUES Or COD_REALTIME
    'End With

End Sub

Public Function ReadControlTEDSSize(bytBoard As Byte, bytChannel As Byte) As Long
    Dim lngResults As Long
    DSS.ReadControlTEDSSize bytBoard, bytChannel, lngResults
    If lngResults = 0 Then
        lngResults = -1
    End If
    ReadControlTEDSSize = lngResults + 1
End Function

Public Function ReadControlTEDS(bytBoard As Byte, bytChannel As Byte) As String
    On Error GoTo ReadControlTEDS_Err

    ' read the control teds
    Dim nSize As Long
    Dim teds As Variant
    Dim strTEDS As String
    Dim intIndex As Integer
    
    If bytChannel = 0 Then
        DSS.ReadControlTEDSSize bytBoard, bytChannel, nSize
        ReDim tedsBuffer(0 To (nSize - 1)) As Byte
        teds = tedsBuffer
    Else
        nSize = 256
        ReDim tedsBuffers(0 To 255) As Byte
        teds = tedsBuffers
    End If
    
    
    If bytChannel = 0 Then
        DSS.ReadControlTEDS bytBoard, bytChannel, teds
    Else
        IReceiver.ReadDSITTEDS bytBoard, bytChannel, teds
    End If
    
    ' Output the TEDS information
    Clipboard.Clear
    strTEDS = "Board " & bytBoard & vbCrLf & "Channel " & bytChannel
    For intIndex = 0 To nSize - 1
        strTEDS = strTEDS & vbCrLf & teds(intIndex)
    Next
    Clipboard.SetText strTEDS
    
    Exit Function

ReadControlTEDS_Err:
    MsgBox Err.Description & vbCrLf & _
           "in MTUProject.modDSS.ReadControlTEDS " & _
           "at line " & Erl
'        Resume Next
End Function

Public Sub ReadSensorTEDS(bytBoard As Byte, bytChannel As Byte)
    Dim vntResults As Variant
    Dim bytResults(0 To 42) As Byte
    Dim strTEDS As String
    Dim intIndex As Integer
    vntResults = bytResults
       
    ' Missing Command
    
    ' Output the TEDS information
    Clipboard.Clear
    strTEDS = "Board " & bytBoard & vbCrLf & "Channel " & bytChannel
    For intIndex = 0 To 42
        strTEDS = strTEDS & vbCrLf & vntResults(intIndex)
    Next
    Clipboard.SetText strTEDS
End Sub


Public Function ReceiverGetSampleRate(bytBoard As Byte, bytChannel As Byte) As Long
    Dim lngResults As Long
    If ChannelSupportsControl(bytBoard, bytChannel, TAG_SAMPLE_RATE) Then
        IReceiver.GetSampleRate bytBoard, bytChannel, lngResults
    Else
        lngResults = -1
    End If
    ReceiverGetSampleRate = lngResults
End Function

Public Function ReceiverGetRange(bytBoard As Byte, bytChannel As Byte) As Long
    Dim lngResults As Long
    If ChannelSupportsControl(bytBoard, bytChannel, TAG_RANGE) Then
        IDualMic.GetRange bytBoard, bytChannel, lngResults
    Else
        lngResults = -1
    End If
    ReceiverGetRange = lngResults
End Function

Public Function ReceiverGetBuffertype(bytBoard As Byte, bytChannel As Byte) As Long
        On Error GoTo ReceiverGetBuffertype_Err
        Dim lngResults As Long
100     If ChannelSupportsControl(bytBoard, bytChannel, TAG_BUFFER_CONTROL) Then
102         IReceiver.GetBufferControl bytBoard, bytChannel, lngResults
        Else
104         lngResults = -1
        End If
        
106     ReceiverGetBuffertype = lngResults

        Exit Function
ReceiverGetBuffertype_Err:
        If Err.Number = 91 Then
            lngResults = -1
            Resume Next
        Else
            MsgBox Err.Description & vbCrLf & _
                   "in Project1.modReceiver.ReceiverGetBuffertype " & _
                   "at line " & Erl
            Resume Next
        End If
End Function


Public Function MicGetSampleRate(bytBoard As Byte, bytChannel As Byte) As Long
    Dim lngResults As Long
    If ChannelSupportsControl(bytBoard, bytChannel, TAG_SAMPLE_RATE) Then
        IDualMic.GetSampleRate bytBoard, bytChannel, lngResults
    Else
        lngResults = -1
    End If
    MicGetSampleRate = lngResults
End Function

Public Sub MicSetSampleRate(bytBoard As Byte, bytChannel As Byte, lngRate As Long)
    IDualMic.SetSampleRate bytBoard, bytChannel, 0
    MsgBox GetDSSError
End Sub

Public Sub MicSetBufferType(bytBoard As Byte, bytChannel As Byte, lngRate As Long)
    IDualMic.SetBufferControl bytBoard, bytChannel, lngRate
End Sub

Public Function MicGetRange(bytBoard As Byte, bytChannel As Byte) As Long
    Dim lngResults As Long
    If ChannelSupportsControl(bytBoard, bytChannel, TAG_RANGE) Then
        IDualMic.GetRange bytBoard, bytChannel, lngResults
    Else
        lngResults = -1
    End If
    MicGetRange = lngResults
End Function

Public Sub MicSetRange(bytBoard As Byte, bytChannel As Byte, lngRange As Long)
    If ChannelSupportsControl(bytBoard, bytChannel, TAG_RANGE) Then
        IDualMic.SetRange bytBoard, bytChannel, lngRange
    End If
End Sub

Public Function MicGetBuffertype(bytBoard As Byte, bytChannel As Byte) As Long
    Dim lngResults As Long
    If ChannelSupportsControl(bytBoard, bytChannel, TAG_BUFFER_CONTROL) Then
        IDualMic.GetBufferControl bytBoard, bytChannel, lngResults
    Else
        lngResults = -1
    End If
    MicGetBuffertype = lngResults
End Function

Public Function MicGetBias(bytBoard As Byte, bytChannel As Byte) As Long
    Dim lngResults As Long
    If ChannelSupportsControl(bytBoard, bytChannel, TAG_BIAS) Then
        IDualMic.GetBias bytBoard, bytChannel, lngResults
    Else
        lngResults = -1
    End If
    MicGetBias = lngResults
End Function

