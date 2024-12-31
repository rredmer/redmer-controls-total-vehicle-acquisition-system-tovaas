VERSION 5.00
Object = "{D8F5B61D-9152-4399-BF30-A1E4F3F072F6}#4.0#0"; "IGTabs40.ocx"
Object = "{85202277-6C76-4228-BC56-7B3E69E8D5CA}#5.0#0"; "IGToolBars50.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmNeoVISetup 
   BorderStyle     =   0  'None
   ClientHeight    =   8220
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrAutoRead 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   30
      Top             =   7770
   End
   Begin ActiveToolBars.SSActiveToolBars MainToolbar 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      ToolBarsCount   =   1
      ToolsCount      =   5
      Tools           =   "frmNeoVIExample.frx":0000
      ToolBars        =   "frmNeoVIExample.frx":3F96
   End
   Begin ActiveTabs.SSActiveTabs MainTab 
      Height          =   7755
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   13679
      _Version        =   262144
      TabCount        =   3
      TagVariant      =   ""
      Tabs            =   "frmNeoVIExample.frx":40C7
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   7365
         Left            =   30
         TabIndex        =   9
         Top             =   360
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   12991
         _Version        =   262144
         TabGuid         =   "frmNeoVIExample.frx":417E
         Begin VB.Frame Frame3 
            Caption         =   "Neo Information"
            Height          =   6495
            Left            =   30
            TabIndex        =   21
            Top             =   60
            Width           =   3735
            Begin VB.CommandButton cmdVersion 
               Caption         =   "ICSNeo40.dll Version"
               Height          =   495
               Left            =   120
               TabIndex        =   25
               Top             =   240
               Width           =   3495
            End
            Begin VB.CommandButton cmdAttachedDevices 
               Caption         =   "Find All Attached Devices"
               Height          =   375
               Left            =   120
               TabIndex        =   24
               Top             =   960
               Width           =   3495
            End
            Begin VB.ListBox lstCommDevices 
               Height          =   2010
               Left            =   120
               TabIndex        =   23
               Top             =   4200
               Width           =   3495
            End
            Begin VB.ListBox lstUsbDevices 
               Height          =   2205
               Left            =   120
               TabIndex        =   22
               Top             =   1680
               Width           =   3495
            End
            Begin VB.Label Label5 
               Caption         =   "Comm Type Devices"
               Height          =   255
               Left            =   120
               TabIndex        =   27
               Top             =   3960
               Width           =   1695
            End
            Begin VB.Label Label3 
               Caption         =   "USB Devices"
               Height          =   255
               Left            =   120
               TabIndex        =   26
               Top             =   1440
               Width           =   1095
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Neo Config Information"
            Height          =   6495
            Left            =   3780
            TabIndex        =   10
            Top             =   60
            Width           =   3255
            Begin VB.CommandButton CmdGetConfiguration 
               Caption         =   "Get Configuration"
               Height          =   375
               Left            =   240
               TabIndex        =   16
               Top             =   240
               Width           =   2775
            End
            Begin VB.ListBox lstConfigInformation 
               Height          =   2985
               Left            =   240
               TabIndex        =   15
               Top             =   600
               Width           =   2775
            End
            Begin VB.CommandButton cmdSendHSCanInfo 
               Caption         =   "Send HS Can Information"
               Height          =   495
               Left            =   960
               TabIndex        =   14
               Top             =   4920
               Width           =   1335
            End
            Begin VB.TextBox txtCNF 
               Height          =   285
               Index           =   0
               Left            =   1680
               TabIndex        =   13
               Text            =   "1"
               Top             =   3840
               Width           =   615
            End
            Begin VB.TextBox txtCNF 
               Height          =   285
               Index           =   1
               Left            =   1680
               TabIndex        =   12
               Text            =   "B8"
               Top             =   4200
               Width           =   615
            End
            Begin VB.TextBox txtCNF 
               Height          =   285
               Index           =   2
               Left            =   1680
               TabIndex        =   11
               Text            =   "5"
               Top             =   4560
               Width           =   615
            End
            Begin VB.Label Label7 
               Caption         =   "CNF1"
               Height          =   255
               Index           =   0
               Left            =   960
               TabIndex        =   20
               Top             =   3840
               Width           =   615
            End
            Begin VB.Label Label7 
               Caption         =   "CNF2"
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   19
               Top             =   4200
               Width           =   615
            End
            Begin VB.Label Label7 
               Caption         =   "CNF3"
               Height          =   255
               Index           =   2
               Left            =   960
               TabIndex        =   18
               Top             =   4560
               Width           =   615
            End
            Begin VB.Label Label9 
               BackColor       =   &H0000C0C0&
               Caption         =   " TIP: use neoVI explorer to get the proper CNFs. ValueCAN CNFs are different than neoVI due to different CAN Chip speeds."
               Height          =   735
               Left            =   120
               TabIndex        =   17
               Top             =   5520
               Width           =   3060
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel ErrorsPanel 
         Height          =   7365
         Left            =   30
         TabIndex        =   2
         Top             =   360
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   12991
         _Version        =   262144
         TabGuid         =   "frmNeoVIExample.frx":41A6
         Begin RichTextLib.RichTextBox ErrorText 
            Height          =   7245
            Left            =   30
            TabIndex        =   8
            Top             =   60
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   12779
            _Version        =   393217
            ReadOnly        =   -1  'True
            ScrollBars      =   3
            TextRTF         =   $"frmNeoVIExample.frx":41CE
         End
      End
      Begin ActiveTabs.SSActiveTabPanel DiagnosticsPanel 
         Height          =   7365
         Left            =   30
         TabIndex        =   1
         Top             =   360
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   12991
         _Version        =   262144
         TabGuid         =   "frmNeoVIExample.frx":4250
         Begin Threed.SSFrame TransmitFrame 
            Height          =   2895
            Left            =   0
            TabIndex        =   5
            Top             =   30
            Width           =   9435
            _ExtentX        =   16642
            _ExtentY        =   5106
            _Version        =   262144
            Caption         =   "Transmit Messages"
            Begin UltraGrid.SSUltraGrid TransmitMessagesGrid 
               Height          =   2145
               Left            =   30
               TabIndex        =   7
               Top             =   210
               Width           =   9345
               _ExtentX        =   16484
               _ExtentY        =   3784
               _Version        =   131072
               GridFlags       =   17040384
               LayoutFlags     =   67108864
               Caption         =   "Transmit Messages"
            End
            Begin VB.CommandButton cmdNeoVITransmit 
               Caption         =   "Transmit"
               Height          =   435
               Left            =   60
               TabIndex        =   6
               Top             =   2400
               Width           =   1335
            End
         End
         Begin Threed.SSFrame ReceiveFrame 
            Height          =   4395
            Left            =   0
            TabIndex        =   3
            Top             =   2910
            Width           =   9435
            _ExtentX        =   16642
            _ExtentY        =   7752
            _Version        =   262144
            Caption         =   "Received Messages"
            Begin RichTextLib.RichTextBox ReceivedMessages 
               Height          =   4095
               Left            =   60
               TabIndex        =   4
               Top             =   210
               Width           =   9285
               _ExtentX        =   16378
               _ExtentY        =   7223
               _Version        =   393217
               ReadOnly        =   -1  'True
               ScrollBars      =   3
               TextRTF         =   $"frmNeoVIExample.frx":4278
            End
         End
      End
   End
End
Attribute VB_Name = "frmNeoVISetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: MTU                                                       **
'**                                                                        **
'** Module.....: frmNeoVISetup                                             **
'**                                                                        **
'** Description: Provides NeoVI Vehicle Bus Interface.                     **
'**                                                                        **
'** History....:                                                           **
'**    12/23/03 v1.71 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2004 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit                                     'Require explicit variable declaration
Private Const MessageLength As Integer = 11         'This is the standard message length (Max)
Private m_hObject As Long                           'Holds the object for the state of the application
Private stMessages(0 To 20000) As icsSpyMessage     'Array of message structures to hold the received data
Public IsConnected As Boolean                       'Set to TRUE when the NeoVI is connected

Private Sub cmdAttachedDevices_Click()
    Dim lResult As Long   ''Holds the Result of the function call
    Dim iDevices(0 To 127) As Long   ''Array containing the
    Dim iSerialNumbers(0 To 127) As Long  ''Array holding the serial numbers of attached devices
    Dim iOpenedStatus(0 To 127) As Long   ''Array holding the Port status of devices
    Dim iNumDevices As Long  ''Tells the number of attached devices
    Dim iCommPortNumbers(0 To 127) As Long  ''Array holding CommPort numbers for each device
    Dim Counter As Integer

        ''Find all USB Device call
        lResult = icsneoFindAllUSBDevices(INTREPIDCS_DRIVER_STANDARD, 1, iDevices(0), iSerialNumbers(0), iOpenedStatus(0), iNumDevices)
        If lResult = 1 Then
            For Counter = 0 To 127
                lstUsbDevices.AddItem ("USB #" & iDevices(Counter) & " SN-" & iSerialNumbers(Counter) & " State-" & iOpenedStatus(Counter))
            Next Counter

        Else
            MsgBox ("Problem Getting Device Information")
        End If
        lResult = icsneoFindAllCOMDevices(INTREPIDCS_DRIVER_STANDARD, 1, 0, 0, iDevices(0), iCommPortNumbers(0), iSerialNumbers(0), iNumDevices)
        If lResult = 1 Then
            For Counter = 0 To 127
                lstCommDevices.AddItem ("Device Type-" & iDevices(Counter) & " SN-" & iSerialNumbers(Counter) & " Port #" & iCommPortNumbers(Counter))
            Next Counter
        Else
                MsgBox ("Could not find anything")
        End If
End Sub


Private Sub CmdGetConfiguration_Click()
    Dim bConfigBytes(1024) As Byte
    Dim lNumBytes As Long
    Dim lResult As Long
    Dim Counter As Integer

    lstConfigInformation.Clear

    lResult = icsneoGetConfiguration(m_hObject, bConfigBytes(0), lNumBytes)
    For Counter = 0 To 1024
        lstConfigInformation.AddItem "Byte Number-" & Counter & " Byte Data-" & bConfigBytes(Counter)
    Next Counter

End Sub

Private Sub cmdSendHSCanInfo_Click()
    Dim bConfigBytes(1024) As Byte   ''Storage for Data bytes from device
    Dim lNumBytes As Long   ''Storage for Number of Bytes
    Dim lResult As Long     ''Storage for Result of Called Function
    Dim Counter As Integer
    
    lstConfigInformation.Clear  ''Clear ListBox
    
    ''Call Get Configuration
    lResult = icsneoGetConfiguration(m_hObject, bConfigBytes(0), lNumBytes)
    
    ''Fill Listbox with Data From Function Call
    For Counter = 0 To 1024
        lstConfigInformation.AddItem "Byte Number-" & Counter & " Byte Data-" & bConfigBytes(Counter)
    Next Counter
    
    ''Set HS CAN Baud Rate information
    bConfigBytes(NEO_CFG_MPIC_HS_CAN_CNF1) = ConvertFromHex(txtCNF(0))
    bConfigBytes(NEO_CFG_MPIC_HS_CAN_CNF2) = ConvertFromHex(txtCNF(1))
    bConfigBytes(NEO_CFG_MPIC_HS_CAN_CNF3) = ConvertFromHex(txtCNF(2))
    
    ''Call Send Configuration
    lResult = icsneoSendConfiguration(m_hObject, bConfigBytes(0), lNumBytes)

    '// make sure the read was successful
    If Not CBool(lResult) Then
        
        MsgBox "Problem sending configuration"
        'cmdCloseNeoVI.Value = True
        Exit Sub
    Else
        MsgBox "Configuration Successfull"
    End If
    

End Sub

Private Sub cmdVersion_Click()
    cmdVersion.Caption = "ICSNeo40.dll Version " & icsneoGetDLLVersion()
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Form_Load                                             **
'**                                                                        **
'**  Description..:  This routine initializes form controls.               **
'**                                                                        **
'****************************************************************************
Private Sub Form_Load()
    ConfigureGrid
    OpenNeoVI
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Form_Unload                                           **
'**                                                                        **
'**  Description..:  This routine closes the NeoVI connection on exit.     **
'**                                                                        **
'****************************************************************************
Private Sub Form_Unload(Cancel As Integer)
    If IsConnected Then CloseNeoVI                  'Close our neoVI object if we opened it
    icsneoFreeObject m_hObject                      'Free the memory associated with our driver object
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Form_Resize                                           **
'**                                                                        **
'**  Description..:  This routine simply resizes controls on form resize.  **
'**                                                                        **
'****************************************************************************
Private Sub Form_Resize()
    If Me.Width > 0 Then
        MainTab.Width = Me.Width
        ReceiveFrame.Width = MainTab.Width - 50
        ReceivedMessages.Width = ReceiveFrame.Width - 200
        ReceivedMessages.Height = ReceiveFrame.Height - 400
        TransmitFrame.Width = MainTab.Width - 50
        TransmitMessagesGrid.Width = TransmitFrame.Width - 200
        ErrorText.Width = MainTab.Width - 200
    End If
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  MainToolBar_ToolClick                                 **
'**                                                                        **
'**  Description..:  This routine handles the toolbar.                     **
'**                                                                        **
'****************************************************************************
Private Sub MainToolBar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "ID_Restart"
            CloseNeoVI
            Sleep 1000
            OpenNeoVI
            MsgBox "NeoVI Restarted.", vbApplicationModal + vbInformation + vbOKOnly, "Notice"
        Case "ID_EraseErrors"
        
        Case "ID_StartPeriodic"
        
        Case "ID_Add"
            If MsgBox("Add new transmit message?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "Add Message") = vbYes Then
                DB.AddTxMessage
            End If
        Case "ID_Erase"
            If MsgBox("Erase transmit message '" & DB.rsTxMessages.Fields("MsgID").Value & "'?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "Erase Message") = vbYes Then
                DB.rsTxMessages.Delete adAffectCurrent
            End If
    End Select
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  tmrAutoRead_Timer                                     **
'**                                                                        **
'**  Description..:  This routine provides timed reading of NeoVI messages.**
'**                                                                        **
'****************************************************************************
Private Sub tmrAutoRead_Timer()
    
    Dim lResult As Long, lNumberOfMessages As Long, lNumberOfErrors As Long, lCount As Long, lByteCount As Long
    Dim sListString As String, dTime As Double
    Dim stJMsg As icsSpyMessageJ1850
    Dim lErrors(0 To 599) As Long
    
    If IsConnected = False Then
        Exit Sub                                    'Do not read messages if we haven't opened neoVI yet
    End If
    
    '---- Read the messages from the driver
    lResult = icsneoGetMessages(m_hObject, stMessages(0), lNumberOfMessages, lNumberOfErrors)
    If CBool(lResult) Then                          'If the read was successful then process messages
        
        'lblReadCount = "Number Read : " & lNumberOfMessages
        
        '----- If there were errors reported by NeoVI - retrieve them & log to Error Tab
        If lNumberOfErrors > 0 Then
            lResult = icsneoGetErrorMessages(m_hObject, lErrors(0), lNumberOfErrors)
            If Not CBool(lResult) Then
                '---- Record fatal error here....
                'MsgBox "Problem Reading Errors"
            Else
                For lCount = 1 To lNumberOfErrors
                    '---- Add to error tab
                    ErrorText.Text = lErrors(lCount - 1) & vbCrLf & ErrorText.Text
                Next
            End If
        End If
        
        '---- Loop for each message read
        For lCount = 1 To lNumberOfMessages
            
            With stMessages(lCount - 1)
            
                '---- Calculate the messages timestamp
                dTime = icsneoGetTimeStamp(.TimeHardware, .TimeHardware2)
                sListString = "Time: " & Format$(dTime, "0.0000") & " "
                
                '---- Was it a tx or rx message
                If (.StatusBitField And icsSpyStatusTx) > 0 Then
                    sListString = sListString & "Tx "
                Else
                    sListString = sListString & "Rx "
                End If
                
                '---- Was it a CAN or other protocol
                Select Case .Protocol
                    Case SPY_PROTOCOL_CAN
                        'list the arb id
                        sListString = sListString & "Network " & GetStringForNetworkID(.NetworkID) & " ArbID : " & Hex(.ArbIDOrHeader) & "  Data "
                    Case Else
                        'list the headers bytes
                        LSet stJMsg = stMessages(lCount - 1)    'copy to a J1850 structure
                        sListString = sListString & "Network " & GetStringForNetworkID(.NetworkID) & " Data : "
                        For lByteCount = 1 To stJMsg.NumberBytesHeader
                            sListString = sListString & Hex(stJMsg.Header(lByteCount)) & " "
                        Next
                End Select
                
                'add the data bytes
                For lByteCount = 1 To .NumberBytesData
                    sListString = sListString & Hex(.data(lByteCount)) & " "
                Next
                
                '---- Add Message to buffer
                ReceivedMessages.Text = sListString & vbCrLf & ReceivedMessages.Text
                
            End With
        Next lCount
    Else
        '---- Log fatal message here
        'MsgBox "Problem Reading Messages"
    End If
    
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  cmdNeoVITransmit_Click                                **
'**                                                                        **
'**  Description..:  This routine transmits NeoVI messages.                **
'**                                                                        **
'****************************************************************************
Private Sub cmdNeoVITransmit_Click()
    
    Dim lResult As Long, lCount As Long, lNetworkID As Long, lNumberBytes As Long
    Dim stMessagesTx As icsSpyMessage
    Dim stJMsg As icsSpyMessageJ1850
    
    If Not IsConnected Then
        MsgBox "neoVI not opened"
        Exit Sub
    End If
    
    If DB.rsTxMessages.RecordCount > 0 Then
        DB.rsTxMessages.MoveFirst
        Do While Not DB.rsTxMessages.EOF
            If DB.rsTxMessages.Fields("Enabled").Value = True Then
                '---- Read the Network we will transmit on (indicated by lstNetwork ListBox)
                lNetworkID = DB.rsTxMessages.Fields("Interface").Value + 1
            
                '---- Is this a CAN network or a J1850/ISO one?
                If lNetworkID <= 4 Then                                 'Its a CAN network
                    '---- Load the message structure
                    With stMessagesTx
                        .ArbIDOrHeader = ConvertFromHex(DB.rsTxMessages.Fields("CAN_ArbID").Value)        'The ArbID
                        .NumberBytesData = DB.rsTxMessages.Fields("DataLength").Value
                        If .NumberBytesData > 8 Then .NumberBytesData = 8   'You can only have 8 databytes with CAN
                        '---- Load all of the data bytes in the structure
                        For lCount = 1 To .NumberBytesData
                            .data(lCount) = ConvertFromHex(DB.rsTxMessages("Byte" & Trim(Str(lCount))).Value)
                        Next
                    End With
                Else                                                    'Not a CAN network
                    '---- Load the message structure (the J1850 struture type)
                    With stJMsg
                        lNumberBytes = DB.rsTxMessages.Fields("DataLength").Value
                        If lNumberBytes > 3 Then                        'how many header (max 3 header bytes) and data bytes
                            .NumberBytesHeader = 3
                            .NumberBytesData = lNumberBytes - 3
                        Else
                            .NumberBytesHeader = lNumberBytes
                            .NumberBytesData = 0
                        End If
                        '---- For all the header bytes
                        For lCount = 1 To .NumberBytesHeader
                            .Header(lCount) = ConvertFromHex(DB.rsTxMessages("Byte" & Trim(Str(lCount))).Value)
                        Next
                        '---- For all the data bytes
                        For lCount = 1 To .NumberBytesData
                            .data(lCount) = ConvertFromHex(DB.rsTxMessages("Byte" & Trim(Str(3 + lCount - 1))).Value)
                        Next
                    End With
                    LSet stMessagesTx = stJMsg                          'Copy the J1850 message structur into the structure that will be transmitted
                End If
                
                '---- Transmit the assembled message
                lResult = icsneoTxMessages(m_hObject, stMessagesTx, lNetworkID, 1)
                If Not CBool(lResult) Then
                    MsgBox "Problem Transmitting Message"
                End If
            End If
            DB.rsTxMessages.MoveNext                                     'Go to the next record
        Loop
    End If
    
End Sub

'The function converts a Hex string and returns a decimal result
Private Function ConvertFromHex(sValue As String) As Double
    ConvertFromHex = Val("&H" & sValue)
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  OpenNeoVI                                             **
'**                                                                        **
'**  Description..:  This routine opens the NeoVI Connection.              **
'**                                                                        **
'****************************************************************************
Private Sub OpenNeoVI()

    Dim lResult As Long, lCount As Long             'General Purpose Counter Variable
    Dim bNetworkIDs(0 To 16) As Byte                'Array of network IDs passed to the driver
    Dim bSCPIDs(0 To 255) As Byte                   'Array of SCP functional IDs passed to the driver
    
    If IsConnected Then Exit Sub                    'Exit if NeoVI is already connected
    
    '---- Initialize the network id array
    For lCount = 0 To 16
        bNetworkIDs(lCount) = lCount
    Next lCount
    'NEOVI_COMMTYPE_USB_BULK
    '---- Open the first neoVi on USB
    lResult = icsneoOpenPort(18, NEOVI_COMMTYPE_RS232, INTREPIDCS_DRIVER_STANDARD, bNetworkIDs(0), bSCPIDs(0), m_hObject)
    If CBool(lResult) Then
        AppLog InfoMsg, "OpenNeoVI, Opened NeoVI successfully."
        IsConnected = True                          'Set the flag which indicates we haved opened neoVI
        tmrAutoRead.Enabled = True
    Else
        AppLog ErrorMsg, "OpenNeoVI, Error Opening NeoVI."
    End If

End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  CloseNeoVI                                            **
'**                                                                        **
'**  Description..:  This routine closes the NeoVI Connection.             **
'**                                                                        **
'****************************************************************************
Private Sub CloseNeoVI()
    Dim lResult As Long, lNumberOfErrors As Long    'Used to store the return variable of the closeport method
    
    tmrAutoRead.Enabled = False                     'Turn off the timer reading the NeoVI Messages
    
    '---- Close the port associated with neoVI
    lResult = icsneoClosePort(m_hObject, lNumberOfErrors)
    If CBool(lResult) Then
        AppLog InfoMsg, "Closed the NeoVI Port."
        IsConnected = False
    Else
        AppLog ErrorMsg, "Error closing the NeoVI Port."
    End If
    
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  CloseNeoVI                                            **
'**                                                                        **
'**  Description..:  This routine returns a description for a network ID.  **
'**                                                                        **
'****************************************************************************
Private Function GetStringForNetworkID(ByVal lNetworkID As Long) As String
    Select Case lNetworkID
        Case NETID_HSCAN
            GetStringForNetworkID = "HSCAN"
        Case NETID_MSCAN
            GetStringForNetworkID = "MSCAN"
        Case NETID_SWCAN
            GetStringForNetworkID = "SWCAN"
        Case NETID_LSFTCAN
            GetStringForNetworkID = "LSFTCAN"
        Case NETID_FORDSCP
            GetStringForNetworkID = "FORD SCP"
        Case NETID_J1708
            GetStringForNetworkID = "J1708"
        Case NETID_AUX
            GetStringForNetworkID = "AUX"
        Case NETID_JVPW
            GetStringForNetworkID = "J1850 VPW"
        Case NETID_ISO
            GetStringForNetworkID = "ISO/UART"
    End Select
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ConfigureGrid                                         **
'**                                                                        **
'**  Description..:  This routine configures the transmit grid.            **
'**                                                                        **
'****************************************************************************
Private Sub ConfigureGrid()
    Dim lCount As Long  'general purpose counter variable
    
    With TransmitMessagesGrid
        Set .DataSource = DB.rsTxMessages
        .Refresh ssRefetchAndFireInitializeRow
        
        .ValueLists.Add "Networks"
        .ValueLists.Add "Lengths"
        
        .ValueLists.Item("Networks").ValueListItems.Add 0, "HSCAN"
        .ValueLists.Item("Networks").ValueListItems.Add 1, "MSCAN"
        .ValueLists.Item("Networks").ValueListItems.Add 2, "SWCAN"
        .ValueLists.Item("Networks").ValueListItems.Add 3, "LFSTCAN"
        .ValueLists.Item("Networks").ValueListItems.Add 4, "Ford SCP"
        .ValueLists.Item("Networks").ValueListItems.Add 5, "J1708"
        .ValueLists.Item("Networks").ValueListItems.Add 6, "Aux Net"
        .ValueLists.Item("Networks").ValueListItems.Add 7, "J1850 PWM"
        .ValueLists.Item("Networks").ValueListItems.Add 8, "ISO"
        
        '---- This is the length of the data message (load the list with 0 to 11)
        For lCount = 0 To MessageLength
            .ValueLists.Item("Lengths").ValueListItems.Add lCount, Str(lCount)
        Next lCount
        
        .Bands(0).Columns(0).Hidden = True              'Hide the project name
        
        .Bands(0).Columns(5).ValueList = "Networks"
        .Bands(0).Columns(5).Style = ssStyleDropDownList
        
        .Bands(0).Columns(7).ValueList = "Lengths"
        .Bands(0).Columns(7).Style = ssStyleDropDownList
        
    End With
    
End Sub
