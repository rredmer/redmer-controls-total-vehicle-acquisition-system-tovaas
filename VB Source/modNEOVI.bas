Attribute VB_Name = "bas_neoVI"
Option Explicit

'//
'// neoVI declares, types, and constants for the WIN32 DLL
'//
'// Copyright 2000-2003 Intrepid Control Systems, Inc.
'// www.intrepidcs.com
'//


'// Message Timestamp neoVI 4
Public Const NEOVI_TIMEHARDWARE2_SCALING = 0.1048576
Public Const NEOVI_TIMEHARDWARE_SCALING = 0.00000169

'// Constants used to calculate the timestamp
Public Const NEOVIPRO_VCAN_TIMEHARDWARE2_SCALING = 0.065536
Public Const NEOVIPRO_VCAN_TIMEHARDWARE_SCALING = 0.000001

'// Configuration Array constants
'// high speed CAN neoVI / valuecan baud rate constants
Public Const NEO_CFG_MPIC_HS_CAN_CNF1 = 512 + 10
Public Const NEO_CFG_MPIC_HS_CAN_CNF2 = 512 + 9
Public Const NEO_CFG_MPIC_HS_CAN_CNF3 = 512 + 8
Public Const NEO_CFG_MPIC_HS_CAN_MODE = 512 + 54

'// med speed CAN
Public Const NEO_CFG_MPIC_MS_CAN_CNF1 = 512 + 22
Public Const NEO_CFG_MPIC_MS_CAN_CNF2 = 512 + 21
Public Const NEO_CFG_MPIC_MS_CAN_CNF3 = 512 + 20

Public Const NEO_CFG_MPIC_SW_CAN_CNF1 = 512 + 34
Public Const NEO_CFG_MPIC_SW_CAN_CNF2 = 512 + 33
Public Const NEO_CFG_MPIC_SW_CAN_CNF3 = 512 + 32

Public Const NEO_CFG_MPIC_LSFT_CAN_CNF1 = 512 + 46
Public Const NEO_CFG_MPIC_LSFT_CAN_CNF2 = 512 + 45
Public Const NEO_CFG_MPIC_LSFT_CAN_CNF3 = 512 + 44

'// Network ID
Public Const NETID_DEVICE = 0
Public Const NETID_HSCAN = 1
Public Const NETID_MSCAN = 2
Public Const NETID_SWCAN = 3
Public Const NETID_LSFTCAN = 43
Public Const NETID_FORDSCP = 5
Public Const NETID_J1708 = 6
Public Const NETID_AUX = 70
Public Const NETID_JVPW = 8
Public Const NETID_ISO = 9
Public Const NETID_ISOPIC = 10
Public Const NETID_MAIN51 = 11
Public Const NETID_HOST = 12

'// Protocols
Public Const SPY_PROTOCOL_CUSTOM = 0
Public Const SPY_PROTOCOL_CAN = 1
Public Const SPY_PROTOCOL_GMLAN = 2
Public Const SPY_PROTOCOL_J1850VPW = 3
Public Const SPY_PROTOCOL_J1850PWM = 4
Public Const SPY_PROTOCOL_ISO9141 = 5
Public Const SPY_PROTOCOL_Keyword2000 = 6
Public Const SPY_PROTOCOL_GM_ALDL_UART = 7
Public Const SPY_PROTOCOL_CHRYSLER_CCD = 8
Public Const SPY_PROTOCOL_CHRYSLER_SCI = 9
Public Const SPY_PROTOCOL_FORD_UBP = 10
Public Const SPY_PROTOCOL_BEAN = 11
Public Const SPY_PROTOCOL_LIN = 12

'// Driver Type Constants
Public Const INTREPIDCS_DRIVER_STANDARD = 0
Public Const INTREPIDCS_DRIVER_TEST = 1

'// Port Type Constants
Public Const NEOVI_COMMTYPE_RS232 = 0
Public Const NEOVI_COMMTYPE_USB_BULK = 1
Public Const NEOVI_COMMTYPE_USB_ISO = 2
Public Const NEOVI_COMMTYPE_TCPIP = 3

'// device Type IDs
Public Const INTREPIDCS_DEVICE_NEO4 = 0
Public Const INTREPIDCS_DEVICE_VCAN = 1
Public Const INTREPIDCS_DEVICE_NEO6 = 2
Public Const INTREPIDCS_DEVICE_UNKNOWN = 3

'// ISO15765 Bit Parameters
Public Enum icsspy15765RxBitfield
    icsspy15765RxErrGlobal = 2 ^ 0
    icsspy15765RxErrCFRX_EXP_FF = 2 ^ 1
    icsspy15765RxErrFCRX_EXP_FF = 2 ^ 2
    icsspy15765RxErrSFRX_EXP_CF = 2 ^ 3
    icsspy15765RxErrFFRX_EXP_CF = 2 ^ 4
    icsspy15765RxErrFCRX_EXP_CF = 2 ^ 5
    icsspy15765RxErrCF_TIME_OUT = 2 ^ 6
    icsspy15765RxComplete = 2 ^ 7
    icsspy15765RxInProgress = 2 ^ 8
    icsspy15765RxErrSeqCntInCF = 2 ^ 9
End Enum


'// these are bitmasks for the status bitfield
Public Enum icsSpyDataStatusBitfield
    icsSpyStatusGlobalError = 2 ^ 0
    icsSpyStatusTx = 2 ^ 1
    icsSpyStatusXtdFrame = 2 ^ 2
    icsSpyStatusRemoteFrame = 2 ^ 3
    icsSpyStatusErrCRCError = 2 ^ 4
    icsSpyStatusCANErrorPassive = 2 ^ 5
    icsSpyStatusErrIncompleteFrame = 2 ^ 6
    icsSpyStatusErrLostArbitration = 2 ^ 7
    icsSpyStatusErrUndefined = 2 ^ 8
    icsSpyStatusErrCANBusOff = 2 ^ 9
    icsSpyStatusErrCANErrorWarning = 2 ^ 10
    icsSpyStatusBusShortedPlus = 2 ^ 11
    icsSpyStatusBusShortedGnd = 2 ^ 12
    icsSpyStatusCheckSumError = 2 ^ 13
    icsSpyStatusErrBadMessageBitTimeError = 2 ^ 14
    icsSpyStatusIFRData = 2 ^ 15
    icsSpyStatusHardwareCommError = 2 ^ 16
    icsSpyStatusExpectedLengthError = 2 ^ 17
    icsSpyStatusIncomingNoMatch = 2 ^ 18
    icsSpyStatusBreak = 2 ^ 19
    icsSpyStatusAVT_VSIRecOverflow = 2 ^ 20
    icsSpyStatusTestTrigger = 2 ^ 21
    icsSpyStatusAudioCommentType = 2 ^ 22
    icsSpyStatusGPSDataValue = 2 ^ 23
    icsSpyStatusAnalogDigitalInputValue = 2 ^ 24
    icsSpyStatusTextCommentType = 2 ^ 25
    icsSpyStatusNetworkMessageType = 2 ^ 26
    icsSpyStatusVSI_TxUnderRun = 2 ^ 27
    icsSpyStatusVSI_IFR_CRCBit = 2 ^ 28
    icsSpyStatusInitMessage = 2 ^ 29
    icsSpyStatusHighSpeed = 2 ^ 30
End Enum

Public Enum icsSpyDataStatusBitfield2
    icsSpyStatusHasValue = 2 ^ 0
    icsSpyStatusValueIsBoolean = 2 ^ 1
    icsSpyStatusHighVoltage = 2 ^ 2
    icsSpyStatusLongMessage = 2 ^ 3
End Enum


'//
Public Type icsSpyMessage
    StatusBitField As Long '4
    StatusBitField2 As Long 'new '4
    TimeHardware As Long ' 4
    TimeHardware2 As Long 'new ' 4
    TimeSystem As Long ' 4
    TimeSystem2 As Long
    TimeStampHardwareID As Byte 'new ' 1
    TimeStampSystemID As Byte
    NetworkID As Byte 'new ' 1
    NodeID As Byte
    Protocol As Byte
    MessagePieceID As Byte ' 1
    ColorID As Byte '1
    NumberBytesHeader As Byte ' 1
    NumberBytesData As Byte ' 1
    DescriptionID As Integer ' 2
    ArbIDOrHeader As Long    '// Holds (up to 3 byte 1850 header or 29 bit CAN header) '4
    Data(1 To 8) As Byte '8
    AckBytes(1 To 8) As Byte 'new '8
    Value As Single ' 4
    MiscData As Byte
End Type

'//
Public Type icsSpyMessageLong
    StatusBitField As Long ' 4
    StatusBitField2 As Long 'new '4
    TimeHardware As Long
    TimeHardware2 As Long 'new ' 4
    TimeSystem As Long '4
    TimeSystem2 As Long
    TimeStampHardwareID As Byte 'new ' 1
    TimeStampSystemID As Byte
    NetworkID As Byte 'new ' 1
    NodeID As Byte
    Protocol As Byte
    MessagePieceID As Byte ' 1
    ColorID As Byte ' 1
    NumberBytesHeader As Byte '
    NumberBytesData As Byte '2
    DescriptionID As Integer '2
    ArbIDOrHeader As Long    '// Holds (up to 3 byte 1850 header or 29 bit CAN header)
    DataMsb As Long
    DataLsb As Long
    AckBytes(1 To 8) As Byte 'new '8
    Value As Single
    MiscData As Byte
    
End Type

'//
Public Type icsSpyMessageJ1850
    StatusBitField As Long '4
    StatusBitField2 As Long 'new '4
    TimeHardware As Long '4
    TimeHardware2 As Long 'new ' 4
    TimeSystem As Long '4
    TimeSystem2 As Long
    TimeStampHardwareID As Byte 'new ' 1
    TimeStampSystemID As Byte
    NetworkID As Byte 'new ' 1
    NodeID As Byte
    Protocol As Byte
    MessagePieceID As Byte ' 1 new
    ColorID As Byte ' 1
    NumberBytesHeader As Byte '1
    NumberBytesData As Byte '1
    DescriptionID As Integer '2
    Header(1 To 4) As Byte  '4  '// Holds (up to 3 byte 1850 header or 29 bit CAN header)
    Data(1 To 8) As Byte '8
    AckBytes(1 To 8) As Byte 'new '8
    Value As Single '4
    MiscData As Byte
End Type


Public Type spyFilterLong
    StatusValue As Long
    StatusMask As Long
    Status2Value As Long
    Status2Mask As Long
    Header As Long
    HeaderMask As Long
    MiscData As Long
    MiscDataMask As Long
    ByteDataMsb As Long
    ByteDataLsb As Long
    ByteDataMaskMsb As Long
    ByteDataMaskLsb As Long
    HeaderLength As Long
    ByteDataLength As Long
    NetworkID As Long
    FrameMaster As Boolean
    bStuff1 As Byte
    bStuff2 As Byte
    ExpectedLength As Long
    NodeID As Long
End Type

    

Public Type spyFilterBytes
    StatusValue As Long
    StatusMask As Long
    Status2Value As Long
    Status2Mask As Long
    Header(1 To 4) As Byte
    HeaderMask(1 To 4) As Byte
    MiscData As Long
    MiscDataMask As Long
    ByteData(1 To 8) As Byte
    ByteDataMask(1 To 8) As Byte
    HeaderLength As Long
    ByteDataLength As Long
    NetworkID As Long
    FrameMaster As Boolean
    bStuff1 As Byte
    bStuff2 As Byte
    ExpectedLength As Long
    NodeID As Long
End Type



'// Function Declares /////////////////////////////////////////////////////////////////////////////
Public Declare Function icsneoOpenPort Lib "icsneo40.dll" (ByVal lPortNumber As Long, _
                            ByVal lPortType As Long, ByVal lDriverType As Long, _
                            ByRef bNetworkID As Byte, ByRef bSCPFunctionID As Byte, ByRef hObject As Long) As Long
Public Declare Function icsneoOpenPortEx Lib "icsneo40.dll" (ByVal lPortNumber As Long, _
                                                            ByVal lPortType As Long, ByVal lDriverID As Long, _
                                                            ByVal lIPAddressMSB As Long, ByVal lIPAddressLSBOrBaudRate As Long, _
                                                            ByVal lForceConfigRead As Long, _
                                                            ByRef bNetworkID As Byte, ByRef hObject As Long) As Long

Public Declare Function icsneoFindAllCOMDevices Lib "icsneo40.dll" (ByVal lDriverID As Long, _
        ByVal lGetSerialNumbers As Long, ByVal lStopAtFirst As Long, _
        ByVal iUSBCommOnly As Long, ByRef p_lDeviceTypes As Long, _
        ByRef p_lComPorts As Long, ByRef p_lSerialNumbers As Long, ByRef lNumDevices As Long) As Long

Public Declare Function icsneoFindAllUSBDevices Lib "icsneo40.dll" _
                (ByVal lDriverID As Long, ByVal lGetSerialNumbers As Long, ByRef p_lDevices As Long, _
                ByRef p_lSerialNumbers As Long, ByRef p_lOpenedDevices As Long, ByRef lNumDevices As Long) As Long

Public Declare Function icsneoClosePort Lib "icsneo40.dll" (ByVal hObject As Long, ByRef pNumberOfErrors As Long) As Long
Public Declare Function icsneoGetMessages Lib "icsneo40.dll" _
                (ByVal hObject As Long, ByRef pMsg As icsSpyMessage, _
                 ByRef pNumberOfMessages As Long, ByRef pNumberOfErrors As Long) As Long
Public Declare Function icsneoTxMessages Lib "icsneo40.dll" (ByVal hObject As Long, _
                                        ByRef pMsg As icsSpyMessage, ByVal lNetwork As Long, ByVal lNumMessages As Long) As Long
Public Declare Sub icsneoFreeObject Lib "icsneo40.dll" (ByVal hObject As Long)
Public Declare Function icsneoGetErrorMessages Lib "icsneo40.dll" _
                (ByVal hObject As Long, ByRef p_lErrorsMsg As Long, ByRef lNumberOfErrors As Long) As Long

Public Declare Sub icsneoGetISO15765Status Lib "icsneo40.dll" (ByVal hObject As Long, _
                    ByVal lNetwork As Long, ByVal lClearTxStatus As Long, _
                    ByVal lClearRxStatus As Long, ByRef lTxStatus As Long, ByRef lRxStatus As Long)

Public Declare Sub icsneoSetISO15765RxParameters Lib "icsneo40.dll" (ByVal hObject As Long, ByVal lNetwork As Long, ByVal lEnable As Long, _
                                                        pFF_CFMsgFilter As spyFilterLong, _
                                                       pFlowCTxMsg As icsSpyMessage, ByVal lCFTimeOutMs As Long, _
                                                       ByVal lFlowCBlockSize As Long, _
                                                       ByVal lUsesExtendedAddressing As Long, ByVal lUseHardwareIfPresent As Long)
Public Declare Function icsneoGetDLLVersion Lib "icsneo40.dll" () As Long
Public Declare Function icsneoStartSockServer Lib "icsneo40.dll" (ByVal hObject As Long, ByVal iPort As Long) As Long
Public Declare Function icsneoStopSockServer Lib "icsneo40.dll" (ByVal hObject As Long) As Long



Private Declare Function icsneoGetErrorInfo Lib "icsneo40.dll" _
                (ByVal lErrorNumber As Long, ByVal sErrorDescriptionShort As String, _
                ByVal sErrorDescriptionLong As String, ByRef lMaxLengthShort As Long, _
                ByRef lMaxLengthLong As Long, ByRef lErrorSeverity As Long, ByRef lRestartNeeded As Long) As Long
                
Public Declare Function icsneoGetConfiguration Lib "icsneo40.dll" _
                (ByVal hObject As Long, ByRef p_bData As Byte, ByRef lNumBytes As Long) As Long
Public Declare Function icsneoSendConfiguration Lib "icsneo40.dll" _
                (ByVal hObject As Long, ByRef p_bData As Byte, ByVal lNumBytes As Long) As Long
                
                
                
'///////////////////////////////////////////////////////////////////////////////////////////////////

Public Function icsneoGetDLLErrorInfo(ByVal lErrorNum As Long, sErrorShort As String, _
                    sErrorLong As String, lSeverity As Long, bRestart As Boolean) As Boolean

    Dim lErrorLongLength As Long
    Dim lErrorShortLength As Long
    Dim lRestart As Long
    Dim lResult As Long
    
    sErrorLong = String(255, 0)
    sErrorShort = String(255, 0)
    
    lErrorLongLength = 255
    lErrorShortLength = 255

    lResult = icsneoGetErrorInfo(lErrorNum, sErrorShort, sErrorLong, _
                                    lErrorShortLength, lErrorLongLength, lSeverity, lRestart)


    sErrorShort = Left$(sErrorShort, lErrorShortLength)
    sErrorLong = Left$(sErrorLong, lErrorLongLength)
    bRestart = CBool(lRestart)
    icsneoGetDLLErrorInfo = CBool(lResult)
    
End Function


'// Timestamp function
Public Function icsneoGetTimeStamp(TimeHardware As Long, TimeHardware2 As Long) As Double
    icsneoGetTimeStamp = NEOVI_TIMEHARDWARE2_SCALING * TimeHardware2 + NEOVI_TIMEHARDWARE_SCALING * TimeHardware
End Function
'// Timestamp function
Public Function icsneoGetTimeStampVCANneoPRO(TimeHardware As Long, TimeHardware2 As Long) As Double
    icsneoGetTimeStampVCANneoPRO = NEOVIPRO_VCAN_TIMEHARDWARE2_SCALING * TimeHardware2 + NEOVIPRO_VCAN_TIMEHARDWARE_SCALING * TimeHardware
End Function


Public Function CreateIPParts(sIPAddress As String, _
                                lIPMsb As Long, _
                                lIPLsb As Long) As Boolean
    Dim vParts As Variant
    Dim dValue As Double
    
    
    vParts = Split(sIPAddress, ".")
     
    If UBound(vParts) <> 3 Then Exit Function
    
    dValue = vParts(3)
    If dValue < 0 Then Exit Function
    If dValue > 255 Then Exit Function
    
    lIPMsb = CLng(dValue) * 256
    
    dValue = vParts(2)
    If dValue < 0 Then Exit Function
    If dValue > 255 Then Exit Function
    
    lIPMsb = lIPMsb + CLng(dValue)
    
    dValue = vParts(1)
    If dValue < 0 Then Exit Function
    If dValue > 255 Then Exit Function
    
    lIPLsb = CLng(dValue) * 256
    
    dValue = vParts(0)
    If dValue < 0 Then Exit Function
    If dValue > 255 Then Exit Function
    
    lIPLsb = lIPLsb + CLng(dValue)
     
     CreateIPParts = True

End Function


