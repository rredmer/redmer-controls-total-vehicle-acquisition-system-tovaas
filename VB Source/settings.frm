VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CANvbas - Settings"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   2400
      Width           =   1455
   End
   Begin VB.ComboBox Hardware 
      Height          =   315
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox Baudrate 
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Text            =   "500000"
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Hardware"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Baudrate"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------
'File:
'  settings.frm
'Project:
'   Defines settings in CANvbas
'-----------------------------------------------------------------------------
' Copyright (c) 1998 by Vector Informatik GmbH.  All rights reserved.
' ----------------------------------------------------------------------------

Private Sub Command1_Click()

Form1.Hide
Main.Show

End Sub

Private Sub Command2_Click()

Unload Form1


End Sub

Private Sub Form_Load()
 
Form1.Hardware.AddItem "None", HWTYPE_NONE
Form1.Hardware.AddItem "Virtual", HWTYPE_VIRTUAL
Form1.Hardware.AddItem "CANcardX", HWTYPE_CANCARDX
Form1.Hardware.AddItem "CANpari", HWTYPE_CANPARI
Form1.Hardware.AddItem "CanDongle", HWTYPE_CANDONGLE
Form1.Hardware.AddItem "CAN-AC2", HWTYPE_CANAC2
Form1.Hardware.AddItem "CAN-AC2-PCI", HWTYPE_CANAC2PCI
Form1.Hardware.AddItem "CanCard", HWTYPE_CANCARD
Form1.Hardware.ListIndex = 2

End Sub

