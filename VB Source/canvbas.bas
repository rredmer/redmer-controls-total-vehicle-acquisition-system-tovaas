Attribute VB_Name = "modVectorCANXL"
'----------------------------------------------------------------------------
'File:
'  CanVbas.bas
'Project:
'   Common functions for the demonstartion application CANvbas
'-----------------------------------------------------------------------------
' Copyright (c) 1998 by Vector Informatik GmbH.  All rights reserved.
' ----------------------------------------------------------------------------

Public Const RX_COUNT = 2

Public vErr As Byte

Public gBaudRate As Long
Public btr0 As Byte
Public btr1 As Byte

Public gPortHandle As Long
Public gChannelMask As Long
Public gOutputMode As Byte
Public gInitMask As Long
Public gPermissionMask As Long
Public gDebugLevel As Byte
Public gTimerRate As Long


Public ev As vbEvent
Public pEvent As vbEvent
' Public rxEv(RX_COUNT) As vbEvent

Public Activated As Boolean
Public count As Long
Public i As Long
Public c As Long
Public chanMask As Long
Public chanIndex As Long
Public idSelect As Long
Public silent As Long
Public messageCount As Long
Public overrunCount As Long
Public lastTime As Double
Public transmited As Boolean
Public transmitCounter As Long

Public resetAcceptance(7) As Boolean
Public h As Long
Public TmpCfg As vbDriverConfig

Public Sub Declarations()

Activated = False
transmited = False
transmitCounter = 0

chanMask = &H1
chanIndex = 0
idSelect = &H1
silent = 0
messageCount = 0
overrunCount = 0
lastTime = 0

gBaudRate = 0
btr0 = &H0
btr1 = &H23

gPortHandle = INVALID_PORTHANDLE
gChannelMask = 0
gOutputMode = OUTPUT_MODE_NORMAL
gInitMask = 0
gPermissionMask = 0
gDebugLevel = 0
gTimerRate = 0

End Sub

Public Function InitDriver()

Dim chipInit As vbChipParams
Dim acc As vbSetAcceptance
gChannelMask = 0

For aa = 0 To 7
  frmVectorCANXL.acceptance_from(aa).Text = ""
  frmVectorCANXL.acceptance_to(aa).Text = ""
  frmVectorCANXL.acceptance_region(aa).Text = "accept all"
  resetAcceptance(aa) = False
Next

vErr = vbOpenDriver()
If vErr Then Fehler
If frmVectorCANXL.moreInfo.Value Then
  frmVectorCANXL.Output.AddItem ">>> Open Driver"
End If

' Print driver config
frmVectorCANXL.CfgUpdate_Click

' Select the wanted channels
gChannelMask = 0
For i = 0 To TmpCfg.channelCount - 1
   If TmpCfg.channel(i).hwType = HWTYPE_CANPARI Or TmpCfg.channel(i).hwType = HWTYPE_CANCARDX Or TmpCfg.channel(i).hwType = HWTYPE_VIRTUAL Or TmpCfg.channel(i).hwType = HWTYPE_CANAC2 Or TmpCfg.channel(i).hwType = HWTYPE_CANAC2PCI Or TmpCfg.channel(i).hwType = HWTYPE_CANCARDY Or TmpCfg.channel(i).hwType = HWTYPE_CANCARDXL Or TmpCfg.channel(i).hwType = HWTYPE_CANCARD2 Or TmpCfg.channel(i).hwType = HWTYPE_EDICCARD Then
     gChannelMask = gChannelMask Or TmpCfg.channel(i).channelMask
   End If
Next
gInitMask = gChannelMask

' Open a port
If frmVectorCANXL.moreInfo.Value Then
  frmVectorCANXL.Output.AddItem ">>> Open port"
End If
vErr = vbOpenPort(gPortHandle, "CANvbas", gChannelMask, gInitMask, gPermissionMask, 1024)
If vErr Then Fehler

If frmVectorCANXL.moreInfo.Value Then
  frmVectorCANXL.Output.AddItem ">>> Porthandle " & gPortHandle
End If
If gPortHandle = INVALID_PORTHANDLE Then Fehler

If frmVectorCANXL.moreInfo.Value Then
  frmVectorCANXL.Output.AddItem ">>> PermissionMask " & gPermissionMask
End If

' If permission to initialize
If gPermissionMask Then
   If vErr Then Fehler
Else
  If frmVectorCANXL.moreInfo.Value Then
    frmVectorCANXL.Output.AddItem ">>> ERROR !!   No init access"
  End If
End If
   
' Set the acceptance filter
' Standard
acc.mask = &H0  ' open all
acc.code = &H0
If frmVectorCANXL.moreInfo.Value Then
  frmVectorCANXL.Output.AddItem ">>> Set standard acceptance filter:  code = " & Hex$(acc.code) & "   mask = " & Hex$(acc.mask) & " -> open all"
End If
vErr = vbSetChannelAcceptance(gPortHandle, gChannelMask, acc)
If vErr Then Fehler
  
' extendend
acc.mask = &H80000000  ' open all
acc.code = &H80000000
If frmVectorCANXL.moreInfo.Value Then
   frmVectorCANXL.Output.AddItem ">>> Set extended acceptance filter:  code = " & Hex$(acc.code) & "   mask = " & Hex$(acc.mask) & " -> open all"
End If
vErr = vbSetChannelAcceptance(gPortHandle, gChannelMask, acc)
If vErr Then Fehler

' put all selected channels on bus
Activated = False
frmVectorCANXL.DeActivate_Click
If frmVectorCANXL.moreInfo.Value Then
  frmVectorCANXL.Output.AddItem "..."
End If

frmVectorCANXL.Timer1.Enabled = True

InitDriver = VSUCCESS

End Function

Public Sub Fehler()

Dim FehlerString As String * 255

nix = vbGetErrorString(vErr, FehlerString)

MsgBox FehlerString, vbOKOnly, "Error !"


End Sub
