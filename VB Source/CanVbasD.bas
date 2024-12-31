Attribute VB_Name = "modVectorCANXLD"
'----------------------------------------------------------------------------
'File:
'  CanVbasD.bas
'Project:
'   Headerfile for using vbCanIF.dll
'   Includes declarations of constants, structures and functions
'   functions to convert C++ union of structure 'tagData' into VB structures
'-----------------------------------------------------------------------------
' Copyright (c) 1998 by Vector Informatik GmbH.  All rights reserved.
' ----------------------------------------------------------------------------

'-----------------------------------
' Constant expressions
'-----------------------------------

' ReceiveMode
Public Const VCAN_WAIT = 0
Public Const VCAN_POLL = 1
Public Const VCAN_NOTIFY = 2

' Errormsg
Public Const VSUCCESS = 0
Public Const VPENDING = 1
Public Const VERR_QUEUE_IS_EMPTY = 10
Public Const VERR_QUEUE_IS_FULL = 11
Public Const VERR_TX_NOT_POSSIBLE = 12
Public Const VERR_NO_LICENSE = 14
Public Const VERR_WRONG_PARAMETER = 101
Public Const VERR_TWICE_REGISTER = 110
Public Const VERR_INVALID_CHAN_INDEX = 111
Public Const VERR_INVALID_ACCESS = 112
Public Const VERR_PORT_IS_OFFLINE = 113
Public Const VERR_CHAN_IS_ONLINE = 116
Public Const VERR_INVALID_PORT = 118
Public Const VERR_HW_NOT_READY = 120
Public Const VERR_CMD_TIMEOUT = 121
Public Const VERR_HW_NOT_PRESENT = 129
Public Const VERR_NOTIFY_ALREADY_ACTIVE = 131
Public Const VERR_CANNOT_OPEN_DRIVER = 201
Public Const VERROR = 255

' Porthandle
Public Const INVALID_PORTHANDLE = -1

' V_RECEIVE_MSG
Public Const MAX_MSG_LEN = 8
Public Const EXT_MSG = &H80000000     ' signs an extended identifier
Public Const MSGFLAG_ERROR_FRAME = &H1                ' Msg is a bus error
Public Const MSGFLAG_OVERRUN = &H2                    ' Msgs following this has been lost
Public Const MSGFLAG_NERR = &H4                       ' NERR active during this msg
Public Const MSGFLAG_WAKEUP = &H8                     ' Msg rcv'd in wakeup mode
Public Const MSGFLAG_REMOTE_FRAME = &H10
Public Const MSGFLAG_RESERVED_1 = &H20                ' Reserved for future usage
Public Const MSGFLAG_TX = &H40                        ' TX acknowledge
Public Const MSGFLAG_TXRQ = &H80                      ' TX request

' V_CHIP_STATE
Public Const CHIPSTATE_BUSOFF = &H1
Public Const CHIPSTATE_ERROR_PASSIVE = &H2
Public Const CHIPSTATE_ERROR_WARNING = &H4
Public Const CHIPSTATE_ERROR_ACTIVE = &H8

' V_TRANSCEIVER
Public Const TRANSCEIVER_EVENT_ERROR = 1
Public Const TRANSCEIVER_EVENT_CHANGED = 2

Public Const TRANSCEIVER_TYPE_NONE = 0
Public Const TRANSCEIVER_TYPE_251 = 1
Public Const TRANSCEIVER_TYPE_252 = 2
Public Const TRANSCEIVER_TYPE_DNOPTO = 3
Public Const TRANSCEIVER_TYPE_W210 = 4
Public Const MAX_TRANSCEIVER_TYPE = 4

Public Const TRANSCEIVER_LINEMODE_NA = 0
Public Const TRANSCEIVER_LINEMODE_TWO_LINE = 1
Public Const TRANSCEIVER_LINEMODE_CAN_H = 2
Public Const TRANSCEIVER_LINEMODE_CAN_L = 3

Public Const TRANSCEIVER_RESNET_NA = 0
Public Const TRANSCEIVER_RESNET_MASTER = 1
Public Const TRANSCEIVER_RESNET_MASTER_STDBY = 2
Public Const TRANSCEIVER_RESNET_SLAVE = 3

' SET_OUTPUT_MODE
Public Const OUTPUT_MODE_SILENT = 0
Public Const OUTPUT_MODE_NORMAL = 1

' Configuration
' Defines for the supported hardware
Public Const HWTYPE_NONE = 0
Public Const HWTYPE_VIRTUAL = 1
Public Const HWTYPE_CANCARDX = 2
Public Const HWTYPE_CANPARI = 3
Public Const HWTYPE_CANDONGLE = 4
Public Const HWTYPE_CANAC2 = 5
Public Const HWTYPE_CANAC2PCI = 6
Public Const HWTYPE_CANCARDY = 12
Public Const HWTYPE_CANCARDXL = 15
Public Const HWTYPE_CANCARD2 = 17
Public Const HWTYPE_EDICCARD = 19

Public Const MAX_HWTYPE = 19


Public Const MAX_CHAN_NAME = 31
Public Const MAX_DRIVER_NAME = 31

Public Const WAIT_TIMEOUT = &H102

Public Enum vbEevent_type
  V_RECEIVE_MSG = 1
  V_CHIP_STATE = 4
  V_CLOCK_OVERFLOW = 5
  V_TRIGGER = 6
  V_TIMER = 8
  V_TRANSCEIVER = 9
  v_TRANSMIT_MSG = 10
End Enum


'-----------------------------------
' Datastructures
'-----------------------------------

Type vbChipParams
    Bitrate As Long
    sjw As Byte
    tseg1 As Byte
    tseg2 As Byte
    sam As Byte
End Type

Type vbChannelConfig
    channelName As String * 32
    hwType As Byte
    hwIndex As Byte
    hwChannel As Byte
    transceiverType As Byte
    ChannelIndex As Byte
    channelMask As Long
    ' Channel
    isOnBus As Byte
    chipParams As vbChipParams
    outputMode As Byte
    flags As Byte
End Type

Type vbDriverConfig
    driverName As String * 32
    driverVersion As Byte
    driverRevision As Byte
    channelCount As Byte
    channel(31) As vbChannelConfig
End Type


Type vbSetAcceptance
    code As Long    ' unsigned long
    mask As Long    ' unsigned long
End Type


' Vevent
Type vbEvent
   tag As Byte
   chanIndex As Byte
   transId As Byte          ' not implemented !
   portHandle As Byte       ' internal use only !
   timestamp As Long        ' originally unsgd. long
   'general: there are problems in Japan with Strings; -> use only Byte-Arry's !!!
   tagData(13) As Byte      ' union structure of _Vmsg, _VchipState, _Vtransceiver
End Type
   
' Vmsg / C-Union Substructure of Vevent
Type vbMsg                       'Length: 14 Bytes
   flags As Byte                '1
   dlc As Byte                  '1
   data(7) As Byte              '8 to avoid problems in japan with strings
   idBytes(3) As Byte           '4: there are problems in Japan with Strings !!!
End Type

' VchipState / C-Union Substructure of Vevent
Type vbChipState                'Length: 3 Bytes
    busStatus As Byte           '1
    txErrorCounter As Byte      '1
    rxErrorCounter As Byte      '1
End Type

' Vtransceiver / C-Union Substructure of Vevent
Type vbTransceiver               'Length 1
    event As Byte
End Type
    
    
Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As String, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
    
 

'-----------------------------------
' Function declarations
'-----------------------------------
'
' Each driver call returns a status value in the range of 0 to 255 (1 Byte).
' If 0 is turned back the driver call was successfull, a value >0 describes
' the reason for failing. See error codes in VERR-constants for details.

Declare Function vbOpenDriver Lib "vbCanIF.dll" Alias "_vbOpenDriver@0" () As Byte
Declare Function vbCloseDriver Lib "vbCanIF.dll" Alias "_vbCloseDriver@0" () As Byte
Declare Function vbGetChannelCount Lib "vbCanIF.dll" Alias "_vbGetChannelCount@0" () As Byte
Declare Function vbGetDriverConfig Lib "vbCanIF.dll" Alias "_vbGetDriverConfig@8" (ByRef chanCount As Byte, ByRef driverConfig As vbDriverConfig) As Byte
Declare Function vbOpenPort Lib "vbCanIF.dll" Alias "_vbOpenPort@24" (ByRef portHandle As Long, ByVal userName As String, ByVal accessMask As Long, ByVal initMask As Long, ByRef permissionMask As Long, ByVal rxQueueSize As Long) As Long
Declare Function vbGetChannelIndex Lib "vbCanIF.dll" Alias "_vbGetChannelIndex@12" (ByVal hwType As Long, ByVal hwIndex As Long, ByVal hwChannel As Long) As Long
Declare Function vbGetChannelMask Lib "vbCanIF.dll" Alias "_vbGetChannelMask@12" (ByVal hwType As Long, ByVal hwIndex As Long, ByVal hwChannel As Long) As Long
Declare Function vbSetChannelMode Lib "vbCanIF.dll" Alias "_vbSetChannelMode@16" (ByVal portHandle As Long, ByVal accessMask As Byte, ByVal tx As Byte, ByVal txrq As Byte) As Byte
Declare Function vbSetChannelParams Lib "vbCanIF.dll" Alias "_vbSetChannelParams@12" (ByVal portHandle As Long, ByVal accessMask As Long, ByRef chipParams As vbChipParams) As Byte
Declare Function vbSetChannelParamsC200 Lib "vbCanIF.dll" Alias "_vbSetChannelParamsC200@16" (ByVal portHandle As Long, ByVal accessMask As Long, ByVal btr0 As Byte, ByVal btr1 As Byte) As Byte
Declare Function vbSetChannelOutput Lib "vbCanIF.dll" Alias "_vbSetChannelOutput@12" (ByVal portHandle As Long, ByVal accessMask As Long, ByVal mode As Byte) As Byte
Declare Function vbSetChannelAcceptance Lib "vbCanIF.dll" Alias "_vbSetChannelAcceptance@12" (ByVal portHandle As Long, ByVal accessMask As Long, ByRef filter As vbSetAcceptance) As Byte
Declare Function vbSetTimerRate Lib "vbCanIF.dll" Alias "_vbSetTimerRate@8" (ByVal portHandle As Long, ByVal timerRate As Long) As Byte
Declare Function vbResetClock Lib "vbCanIF.dll" Alias "_vbResetClock@4" (ByVal portHandel As Long) As Byte
Declare Function vbSetNotification Lib "vbCanIF.dll" Alias "_vbSetNotification@12" (ByVal portHandle As Long, ByRef handle As Long, ByVal queueLevel As Long) As Byte
Declare Function vbActivateChannel Lib "vbCanIF.dll" Alias "_vbActivateChannel@8" (ByVal portHandle As Long, ByVal accessMask As Long) As Byte
Declare Function vbTransmit Lib "vbCanIF.dll" Alias "_vbTransmit@12" (ByVal portHandle As Long, ByVal accessMask As Long, ByRef ev As vbEvent) As Byte
Declare Function vbReceive Lib "vbCanIF.dll" Alias "_vbReceive@20" (ByVal portHandle As Long, ByVal ReceiveMode As Long, ByVal waitHandle As Long, ByRef eventCount As Long, ByRef eventList As vbEvent) As Byte
Declare Function vbReceiveNotify Lib "vbCanIF.dll" Alias "_vbReceive@20" (ByVal portHandle As Long, ByVal ReceiveMode As Long, ByVal waitHandle As Long, ByRef eventCount As Long, ByVal adr As String) As Byte
Declare Function vbReceive1 Lib "vbCanIF.dll" Alias "_vbReceive1@8" (ByVal portHandle As Long, ByRef ev As vbEvent) As Byte
Declare Function vbRequestChipState Lib "vbCanIF.dll" Alias "_vbRequestChipState@8" (ByVal portHandle As Long, ByVal accessMask As Long) As Byte
Declare Function vbFlushTransmitQueue Lib "vbCanIF.dll" Alias "_vbFlushTransmitQueue@8" (ByVal portHandle As Long, ByVal accessMask As Long) As Byte
Declare Function vbFlushReceiveQueue Lib "vbCanIF.dll" Alias "_vbFlushReceiveQueue@4" (ByVal portHandle As Long) As Byte
Declare Function vbGetReceiveQueueLevel Lib "vbCanIF.dll" Alias "_vbGetReceiveQueueLevel@8" (ByVal portHandle As Long, ByRef level As Long) As Byte
Declare Function vbGetState Lib "vbCanIF.dll" Alias "_vbGetState@4" (ByVal portHandle As Long) As Byte
Declare Function vbDeactivateChannel Lib "vbCanIF.dll" Alias "_vbDeactivateChannel@8" (ByVal portHandle As Long, ByVal accessMask As Long) As Byte
Declare Function vbClosePort Lib "vbCanIF.dll" Alias "_vbClosePort@4" (ByVal portHandle As Long) As Byte
Declare Function vbGetErrorString Lib "vbCanIF.dll" Alias "_vbGetErrorString@8" (ByVal errCode As Byte, ByVal errStr As String) As Byte
Declare Function vbGetEventString Lib "vbCanIF.dll" Alias "_vbGetEventString@8" (ByRef ev As vbEvent, ByVal evstr As String) As Byte
Declare Function vbSetChannelBitrate Lib "vbCanIF.dll" Alias "_vbSetChannelBitrate@12" (ByVal portHandle As Long, ByVal accessMask As Long, ByVal Bitrate As Long) As Byte

Declare Function vbGetApplConfig Lib "vbCanIF.dll" Alias "_vbGetApplConfig@20" (ByVal appName As String, ByVal appChannel As Long, ByRef hwType As Long, ByRef hwIndex As Long, ByRef hwChannel As Long) As Byte
Declare Function vbSetApplConfig Lib "vbCanIF.dll" Alias "_vbSetApplConfig@20" (ByVal appName As String, ByVal appChannel As Long, ByVal hwType As Long, ByVal hwIndex As Long, ByVal hwChannel As Long) As Byte
Declare Function vbGetChannelVersion Lib "vbCanIF.dll" Alias "_vbGetChannelVersion@16" (ByVal ChannelIndex As Long, ByRef FwVersion As Long, ByRef HwVersion As Long, ByRef SerialNumber As Long) As Byte
Declare Function vbSetReceiveMode Lib "vbCanIF.dll" Alias "_vbSetReceiveMode@12" (ByVal Port As Long, ByVal ErrorFrame As Byte, ByVal ChipState As Byte) As Byte

' new funktions since version 2.06
Declare Function vbSetChannelTransceiver Lib "vbCanIF.dll" Alias "_vbSetChannelTransceiver@20" (ByVal portHandle As Long, ByVal accessMask As Long, ByVal typ As Long, ByVal lineMode As Long, ByVal resNet As Long) As Byte
' new funktions since version 3.01
Declare Function vbAddAcceptanceRange Lib "vbCanIF.dll" Alias "_vbAddAcceptanceRange@16" (ByVal portHandle As Long, ByVal accessMask As Long, ByVal first_id As Long, ByVal last_id As Long) As Byte
Declare Function vbRemoveAcceptanceRange Lib "vbCanIF.dll" Alias "_vbRemoveAcceptanceRange@16" (ByVal portHandle As Long, ByVal accessMask As Long, ByVal first_id As Long, ByVal last_id As Long) As Byte
Declare Function vbResetAcceptance Lib "vbCanIF.dll" Alias "_vbResetAcceptance@12" (ByVal portHandle As Long, ByVal accessMask As Long, ByVal extended As Long) As Byte


Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
' for internal use only
Public Function Get_vbEvent_tagData_vbMsg(struct As vbEvent) As vbMsg
' returns the substructure 'msg', in C declared as union-structue, of Vevent
' 30.07.98/as.

Dim Tmp As vbMsg

    For i = 0 To 3
      Tmp.idBytes(i) = struct.tagData(i)
    Next i
    Tmp.flags = struct.tagData(4)
    Tmp.dlc = struct.tagData(5)
    For i = 0 To 7
      Tmp.data(i) = struct.tagData(i + 6)
    Next i
    
Get_vbEvent_tagData_vbMsg = Tmp
 
End Function

Public Function Get_vbEvent_tagData_vbChipState(struct As vbEvent) As vbChipState
' returns the substructure 'chipState', in C declared as union-structue, of Vevent
' 30.07.98/as.

Dim Tmp As vbChipState
    
    Tmp.busStatus = struct.tagData(0)
    Tmp.txErrorCounter = struct.tagData(1)
    Tmp.rxErrorCounter = struct.tagData(2)
    
Get_vbEvent_tagData_vbChipState = Tmp

End Function
Public Function Get_vbEvent_tagData_vbTransceiver(struct As vbEvent) As vbTransceiver
' returns the substructure 'transceiver', in C declared as union-structue, of Vevent
' 30.07.98/as.

Dim Tmp As vbTransceiver
    
    Tmp.event = struct.tagData(0)
    
Get_vbEvent_tagData_vbTransceiver = Tmp

End Function

Public Function Build_vbEvent_tagData_vbMsg(TmpEvent As vbEvent, TmpMsg As vbMsg) As vbEvent
' joins a Vmsg structure (TmpMsg) to Vevent.tagData (TmpEvent)
' 05.08.98/as.

For i = 0 To 3
  TmpEvent.tagData(i) = TmpMsg.idBytes(i)
Next i
TmpEvent.tagData(4) = TmpMsg.flags
TmpEvent.tagData(5) = TmpMsg.dlc
For i = 6 To 13
  TmpEvent.tagData(i) = TmpMsg.data(i - 6)
Next i

Build_vbEvent_tagData_vbMsg = TmpEvent
   
End Function
Public Function Build_vbEvent_tagData_vbChipState(TmpEvent As vbEvent, TmpState As vbChipState) As vbEvent
' joins a VchipState structure (TmpState) to Vevent.tagData (TmpEvent)
' 05.08.98/as.

TmpEvent.tagData(0) = TmpState.busStatus
TmpEvent.tagData(1) = TmpState.txErrorCounter
TmpEvent.tagData(2) = TmpState.rxErrorCounter

Build_vbEvent_tagData_vbChipState = TmpEvent

End Function
Public Function Build_vbEvent_tagData_vbTransceiver(TmpEvent As vbEvent, TmpTransceiver As vbTransceiver) As vbEvent
' joins a Vtransceiver structure (TmpTransceiver) to Vevent.tagData (TmpEvent)
' 05.08.98/as.

TmpEvent.tagData(0) = TmpTransceiver.event

Build_vbEvent_tagData_vbTransceiver = TmpEvent
   
End Function

Public Sub LongToByteArray(InNo As Long, OutArray() As Byte)
  Dim f As Long, i As Integer
  
  f = 0
  For i = 0 To 7
    f = f + 2 ^ i
  Next i
  OutArray(0) = InNo And f
  f = 0
  For i = 8 To 15
    f = f + 2 ^ i
  Next i
  OutArray(1) = (InNo And f) / 2 ^ 8
  f = 0
  For i = 16 To 23
    f = f + 2 ^ i
  Next i
  OutArray(2) = (InNo And f) / 2 ^ 16
  f = 0
  For i = 24 To 30
    f = f + 2 ^ i
  Next i
  OutArray(3) = (InNo And f) / 2 ^ 24
End Sub

Public Sub ByteArrayToLong(InArray() As Byte, OutNo As Long)
  OutNo = InArray(0) + InArray(1) * 2 ^ 8 + InArray(2) * 2 ^ 16 + InArray(3) * 2 ^ 24
End Sub


