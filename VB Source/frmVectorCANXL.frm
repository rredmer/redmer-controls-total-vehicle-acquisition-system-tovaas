VERSION 5.00
Begin VB.Form frmVectorCANXL 
   BorderStyle     =   0  'None
   Caption         =   "CANvbas"
   ClientHeight    =   9705
   ClientLeft      =   -180
   ClientTop       =   -285
   ClientWidth     =   10545
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9705
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton channel 
      Caption         =   "no hardware found"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   116
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton acceptance_reset 
      Caption         =   "Reset"
      Height          =   255
      Index           =   7
      Left            =   7320
      TabIndex        =   104
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton acceptance_reset 
      Caption         =   "Reset"
      Height          =   255
      Index           =   6
      Left            =   7320
      TabIndex        =   103
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton acceptance_reset 
      Caption         =   "Reset"
      Height          =   255
      Index           =   5
      Left            =   7320
      TabIndex        =   102
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton acceptance_reset 
      Caption         =   "Reset"
      Height          =   255
      Index           =   4
      Left            =   7320
      TabIndex        =   101
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton acceptance_reset 
      Caption         =   "Reset"
      Height          =   255
      Index           =   3
      Left            =   7320
      TabIndex        =   100
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton acceptance_reset 
      Caption         =   "Reset"
      Height          =   255
      Index           =   2
      Left            =   7320
      TabIndex        =   99
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton acceptance_reset 
      Caption         =   "Reset"
      Height          =   255
      Index           =   1
      Left            =   7320
      TabIndex        =   98
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton acceptance_add 
      Caption         =   "Add"
      Height          =   255
      Index           =   7
      Left            =   6600
      TabIndex        =   97
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton acceptance_add 
      Caption         =   "Add"
      Height          =   255
      Index           =   6
      Left            =   6600
      TabIndex        =   96
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton acceptance_add 
      Caption         =   "Add"
      Height          =   255
      Index           =   5
      Left            =   6600
      TabIndex        =   95
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton acceptance_add 
      Caption         =   "Add"
      Height          =   255
      Index           =   4
      Left            =   6600
      TabIndex        =   94
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton acceptance_add 
      Caption         =   "Add"
      Height          =   255
      Index           =   3
      Left            =   6600
      TabIndex        =   93
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton acceptance_add 
      Caption         =   "Add"
      Height          =   255
      Index           =   2
      Left            =   6600
      TabIndex        =   92
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton acceptance_add 
      Caption         =   "Add"
      Height          =   255
      Index           =   1
      Left            =   6600
      TabIndex        =   91
      Top             =   960
      Width           =   615
   End
   Begin VB.ComboBox acceptance_region 
      Height          =   315
      Index           =   7
      Left            =   5160
      TabIndex        =   90
      Text            =   "accept all"
      Top             =   2400
      Width           =   1335
   End
   Begin VB.ComboBox acceptance_region 
      Height          =   315
      Index           =   6
      Left            =   5160
      TabIndex        =   89
      Text            =   "accept all"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.ComboBox acceptance_region 
      Height          =   315
      Index           =   5
      Left            =   5160
      TabIndex        =   88
      Text            =   "accept all"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ComboBox acceptance_region 
      Height          =   315
      Index           =   4
      Left            =   5160
      TabIndex        =   87
      Text            =   "accept all"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ComboBox acceptance_region 
      Height          =   315
      Index           =   3
      Left            =   5160
      TabIndex        =   86
      Text            =   "accept all"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ComboBox acceptance_region 
      Height          =   315
      Index           =   2
      Left            =   5160
      TabIndex        =   85
      Text            =   "accept all"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.ComboBox acceptance_region 
      Height          =   315
      Index           =   1
      Left            =   5160
      TabIndex        =   84
      Text            =   "accept all"
      Top             =   960
      Width           =   1335
   End
   Begin VB.ListBox Output 
      Height          =   5910
      Left            =   3960
      TabIndex        =   54
      Top             =   3000
      Width           =   6495
   End
   Begin VB.CheckBox moreInfo 
      Caption         =   "Show detail information"
      Height          =   375
      Left            =   2280
      TabIndex        =   43
      Top             =   9120
      Width           =   2055
   End
   Begin VB.CommandButton CANhardware 
      Caption         =   "CAN-Hardware"
      Height          =   375
      Left            =   6720
      TabIndex        =   42
      Top             =   9120
      Width           =   1575
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Number of events"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton Command15 
      Caption         =   "FlushReceiveQueue"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      Caption         =   "FlushTransmitQueue"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Clear Screen"
      Height          =   375
      Left            =   4680
      TabIndex        =   16
      Top             =   9120
      Width           =   1575
   End
   Begin VB.CommandButton OnOffline 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2160
      Top             =   240
   End
   Begin VB.CommandButton Command12 
      Caption         =   "SetNotification"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Toggle Timer"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Toggle &output mode"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Reset clock"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton DeActivate 
      Height          =   375
      Left            =   8040
      TabIndex        =   4
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Re&quest Chipstate"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "remote &message"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "a message &burst"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   " a message"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   4920
      Width           =   3495
   End
   Begin VB.Frame Message 
      Caption         =   "Message"
      Height          =   1575
      Left            =   120
      TabIndex        =   20
      Top             =   3000
      Width           =   3735
      Begin VB.CheckBox extended 
         Caption         =   "extended"
         Height          =   255
         Left            =   2040
         TabIndex        =   118
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox data 
         Height          =   285
         Index           =   7
         Left            =   2880
         MaxLength       =   2
         TabIndex        =   61
         Text            =   "8"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox data 
         Height          =   285
         Index           =   6
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   60
         Text            =   "7"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox data 
         Height          =   285
         Index           =   5
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   59
         Text            =   "6"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox data 
         Height          =   285
         Index           =   4
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   58
         Text            =   "5"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox data 
         Height          =   285
         Index           =   3
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   57
         Text            =   "4"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox data 
         Height          =   285
         Index           =   2
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   56
         Text            =   "3"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox data 
         Height          =   285
         Index           =   1
         Left            =   720
         MaxLength       =   2
         TabIndex        =   55
         Text            =   "2"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox data 
         Height          =   285
         Index           =   0
         Left            =   360
         MaxLength       =   2
         TabIndex        =   31
         Text            =   "1"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox dlc 
         Height          =   315
         Left            =   3120
         MaxLength       =   1
         TabIndex        =   30
         Text            =   "8"
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton TransmitID_Select 
         Caption         =   "0"
         Height          =   315
         Index           =   2
         Left            =   1800
         TabIndex        =   24
         Top             =   480
         Width           =   195
      End
      Begin VB.CommandButton TransmitID_Select 
         Caption         =   ">"
         Height          =   315
         Index           =   1
         Left            =   1560
         TabIndex        =   23
         Top             =   480
         Width           =   195
      End
      Begin VB.TextBox idSelect 
         Height          =   285
         Left            =   360
         MaxLength       =   8
         TabIndex        =   22
         Text            =   "1"
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton TransmitID_Select 
         Caption         =   "<"
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   21
         Top             =   480
         Width           =   195
      End
      Begin VB.Label Label11 
         Caption         =   "ID"
         Height          =   255
         Left            =   480
         TabIndex        =   41
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "7"
         Height          =   255
         Left            =   3000
         TabIndex        =   40
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label9 
         Caption         =   "6"
         Height          =   255
         Left            =   2640
         TabIndex        =   39
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label8 
         Caption         =   "5"
         Height          =   255
         Left            =   2280
         TabIndex        =   38
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label7 
         Caption         =   "4"
         Height          =   255
         Left            =   1920
         TabIndex        =   37
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label6 
         Caption         =   "3"
         Height          =   255
         Left            =   1560
         TabIndex        =   36
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label5 
         Caption         =   "2"
         Height          =   255
         Left            =   1200
         TabIndex        =   35
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label4 
         Caption         =   "1"
         Height          =   255
         Left            =   840
         TabIndex        =   34
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "0"
         Height          =   255
         Left            =   480
         TabIndex        =   33
         Top             =   840
         Width           =   135
      End
      Begin VB.Label DLC_text 
         Caption         =   "DLC"
         Height          =   255
         Left            =   3120
         TabIndex        =   32
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Transmit 
      Caption         =   "Transmit"
      Height          =   1215
      Left            =   120
      TabIndex        =   67
      Top             =   4680
      Width           =   3735
   End
   Begin VB.Frame Frame4 
      Caption         =   "Automatic receive"
      Height          =   615
      Left            =   2040
      TabIndex        =   13
      Top             =   6120
      Width           =   1695
      Begin VB.OptionButton ReceiveMode 
         Caption         =   "No"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton ReceiveMode 
         Caption         =   "Yes"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Counter"
      Height          =   975
      Left            =   8040
      TabIndex        =   9
      Top             =   600
      Width           =   2295
      Begin VB.CommandButton Command14 
         Caption         =   "Reset"
         Height          =   315
         Left            =   1560
         TabIndex        =   12
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Counter 
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Hardware 
      Caption         =   "Hardware"
      Height          =   2775
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   10335
      Begin VB.CheckBox requestTransmit 
         Caption         =   "show transmit acknowledge"
         Enabled         =   0   'False
         Height          =   255
         Left            =   7920
         TabIndex        =   119
         Top             =   2280
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.OptionButton channel 
         Caption         =   "no hardware found"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   117
         Top             =   2280
         Width           =   2655
      End
      Begin VB.OptionButton channel 
         Caption         =   "no hardware found"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   115
         Top             =   1800
         Width           =   2655
      End
      Begin VB.OptionButton channel 
         Caption         =   "no hardware found"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   114
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox acceptance_to 
         Height          =   285
         Index           =   7
         Left            =   4560
         MaxLength       =   3
         TabIndex        =   83
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox acceptance_to 
         Height          =   285
         Index           =   6
         Left            =   4560
         MaxLength       =   3
         TabIndex        =   82
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox acceptance_to 
         Height          =   285
         Index           =   5
         Left            =   4560
         MaxLength       =   3
         TabIndex        =   81
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox acceptance_to 
         Height          =   285
         Index           =   4
         Left            =   4560
         MaxLength       =   3
         TabIndex        =   80
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox acceptance_to 
         Height          =   285
         Index           =   3
         Left            =   4560
         MaxLength       =   3
         TabIndex        =   79
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox acceptance_to 
         Height          =   285
         Index           =   2
         Left            =   4560
         MaxLength       =   3
         TabIndex        =   78
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox acceptance_to 
         Height          =   285
         Index           =   1
         Left            =   4560
         MaxLength       =   3
         TabIndex        =   77
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox acceptance_from 
         Height          =   285
         Index           =   7
         Left            =   3840
         MaxLength       =   3
         TabIndex        =   76
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox acceptance_from 
         Height          =   285
         Index           =   6
         Left            =   3840
         MaxLength       =   3
         TabIndex        =   75
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox acceptance_from 
         Height          =   285
         Index           =   5
         Left            =   3840
         MaxLength       =   3
         TabIndex        =   74
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox acceptance_from 
         Height          =   285
         Index           =   4
         Left            =   3840
         MaxLength       =   3
         TabIndex        =   73
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox acceptance_from 
         Height          =   285
         Index           =   3
         Left            =   3840
         MaxLength       =   3
         TabIndex        =   72
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox acceptance_from 
         Height          =   285
         Index           =   2
         Left            =   3840
         MaxLength       =   3
         TabIndex        =   71
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox acceptance_from 
         Height          =   285
         Index           =   1
         Left            =   3840
         MaxLength       =   3
         TabIndex        =   70
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton acceptance_add 
         Caption         =   "Add"
         Height          =   255
         Index           =   0
         Left            =   6480
         TabIndex        =   66
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton acceptance_reset 
         Caption         =   "Reset"
         Height          =   255
         Index           =   0
         Left            =   7200
         TabIndex        =   65
         Top             =   600
         Width           =   615
      End
      Begin VB.ComboBox acceptance_region 
         Height          =   315
         Index           =   0
         Left            =   5040
         TabIndex        =   64
         Text            =   "accept all"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox acceptance_to 
         Height          =   285
         Index           =   0
         Left            =   4560
         MaxLength       =   3
         TabIndex        =   63
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox acceptance_from 
         Height          =   285
         Index           =   0
         Left            =   3840
         MaxLength       =   3
         TabIndex        =   62
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Bitrate_Ch 
         Height          =   285
         Index           =   7
         Left            =   2880
         TabIndex        =   53
         Text            =   "no bitrate"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox Bitrate_Ch 
         Height          =   285
         Index           =   6
         Left            =   2880
         TabIndex        =   52
         Text            =   "no bitrate"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Bitrate_Ch 
         Height          =   285
         Index           =   5
         Left            =   2880
         TabIndex        =   51
         Text            =   "no bitrate"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Bitrate_Ch 
         Height          =   285
         Index           =   4
         Left            =   2880
         TabIndex        =   50
         Text            =   "no bitrate"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Bitrate_Ch 
         Height          =   285
         Index           =   3
         Left            =   2880
         TabIndex        =   49
         Text            =   "no bitrate"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Bitrate_Ch 
         Height          =   285
         Index           =   2
         Left            =   2880
         TabIndex        =   48
         Text            =   "no bitrate"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Bitrate_Ch 
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   47
         Text            =   "no bitrate"
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton channel 
         Caption         =   "no hardware found"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   46
         Top             =   1320
         Width           =   2655
      End
      Begin VB.OptionButton channel 
         Caption         =   "no hardware found"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   45
         Top             =   1080
         Width           =   2655
      End
      Begin VB.OptionButton channel 
         Caption         =   "no hardware found"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   840
         Width           =   2655
      End
      Begin VB.OptionButton channel 
         Caption         =   "no hardware found"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox Bitrate_Ch 
         Height          =   285
         Index           =   0
         Left            =   2880
         TabIndex        =   29
         Text            =   "no bitrate"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "to"
         Height          =   255
         Index           =   8
         Left            =   4320
         TabIndex        =   112
         Top             =   2280
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "to"
         Height          =   255
         Index           =   7
         Left            =   4320
         TabIndex        =   111
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "to"
         Height          =   255
         Index           =   6
         Left            =   4320
         TabIndex        =   110
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "to"
         Height          =   255
         Index           =   5
         Left            =   4320
         TabIndex        =   109
         Top             =   1560
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "to"
         Height          =   255
         Index           =   4
         Left            =   4320
         TabIndex        =   108
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "to"
         Height          =   255
         Index           =   3
         Left            =   4320
         TabIndex        =   107
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "to"
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   106
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "to"
         Height          =   255
         Index           =   0
         Left            =   4320
         TabIndex        =   69
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Acceptance 
         Caption         =   "Standard acceptance filter "
         Height          =   255
         Left            =   3840
         TabIndex        =   68
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Bitrate 
         Caption         =   "Bitrate"
         Height          =   255
         Left            =   2880
         TabIndex        =   28
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Channels 
         Caption         =   "Channels"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   113
      Top             =   8880
      Width           =   10335
   End
   Begin VB.Label Label3 
      Caption         =   "to"
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   105
      Top             =   960
      Width           =   135
   End
End
Attribute VB_Name = "frmVectorCANXL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------
'File:
'  me.frm
'Project:
'   functions for the main window in CANvbas
'-----------------------------------------------------------------------------
' Copyright (c) 1998 by Vector Informatik GmbH.  All rights reserved.
' ----------------------------------------------------------------------------

Private Sub acceptance_add_Click(Index As Integer)
Dim first_id
Dim last_id
Dim channelMask
channelMask = (gChannelMask And (2 ^ Index))
first_id = Val("&h" & Me.acceptance_from(Index).Text)
last_id = Val("&h" & Me.acceptance_to(Index).Text)
vErr = vbAddAcceptanceRange(gPortHandle, channelMask, first_id, last_id)
If (vErr = 0 And (first_id <> 0) And (last_id <> 0)) Then
  If Me.moreInfo.Value Then
    Me.Output.AddItem ">>> Add acceptance range"
  End If
  Me.acceptance_region(Index).AddItem (Me.acceptance_from(Index).Text & " to " & Me.acceptance_to(Index).Text)
  Me.acceptance_region(Index).Text = (Me.acceptance_from(Index).Text & " to " & Me.acceptance_to(Index).Text)
End If
End Sub

Private Sub acceptance_reset_Click(Index As Integer)
Dim extended
Dim channelMask
channelMask = (gChannelMask And (2 ^ Index))
extended = 0 ' extended is not supported in the driver
vErr = vbResetAcceptance(gPortHandle, channelMask, extended)
If vErr = 0 Then
  If Me.moreInfo.Value Then
    Me.Output.AddItem ">>> Reset acceptance"
  End If
  Me.acceptance_region(Index).Clear
  Me.acceptance_region(Index).Text = "close all"
  Me.acceptance_add(Index).Enabled = True
  resetAcceptance(Index) = True
End If
End Sub
Private Sub CANhardware_Click()
Dim RetVal
RetVal = Shell("vcanconf.EXE", 1)    ' Run Calculator.
End Sub

Private Sub channel_Click(Index As Integer)
chanIndex = Index
chanMask = TmpCfg.channel(chanIndex).channelMask
End Sub

Private Sub Command15_Click()
' Flush receive queue

vErr = vbFlushReceiveQueue(gPortHandle)
If vErr Then
   Fehler
Else
   If Me.moreInfo.Value Then
      Me.Output.AddItem ">>> FlushReceiveQueue"
   End If
End If


End Sub

Public Sub CfgUpdate_Click()
    
Dim n As Byte
Dim channelName As String
Dim channelName1 As String
Dim channelName2 As String
Dim foundSelection As Boolean

foundSelection = False

If vbGetChannelCount < 9 Then

  vErr = vbGetDriverConfig(n, TmpCfg)

  If Me.moreInfo.Value Then
    Me.Output.AddItem ">>> Get driver configuration"
  End If

  If vErr Then Fehler

  'default is false
  For aa = 0 To 7
    Me.channel(aa).Enabled = False
  Next

  'enabel the founded
  For tmpChan = 0 To TmpCfg.channelCount - 1
    channelName = TmpCfg.channel(tmpChan).channelName
    channelName1 = "(c=" & tmpChan & ") " & TmpCfg.channel(tmpChan).channelName
    Me.channel(tmpChan).Caption = channelName1
    Me.Bitrate_Ch(tmpChan).Text = TmpCfg.channel(tmpChan).chipParams.Bitrate
    Me.channel(tmpChan).Enabled = True
    If Me.channel(tmpChan).Value Then
      foundSelection = True
    End If
  Next
  
  ' select the first hardware found if not selected before
  If foundSelection Then
  Else
    For tmpChan = 0 To TmpCfg.channelCount - 1
      If TmpCfg.channel(tmpChan).hwType <> HWTYPE_VIRTUAL Then
        Me.channel(tmpChan).Value = True
        Exit For
      End If
    Next
  End If
Else
    Me.Output.AddItem ">>> only 8 channels are supported in this sample !!!"
End If
End Sub

Private Sub Command1_Click()
'  Transmit a message

Dim TmpMsg As vbMsg
Dim j As Integer
Dim subStr As String
Dim strLen As Integer

Dim transmitID As String
transmitID = Hex$(Val("&h" & Me.idSelect.Text) And &H7FFFFFFF)
If (Me.extended.Value) Then
  transmitID = Hex$(Val("&h" & Me.idSelect.Text) Or &H80000000)
End If

ev.tag = v_TRANSMIT_MSG

j = 0
strLen = Len(transmitID)
For ii = strLen - 1 To 1 Step -2
  subStr = "&h" & Mid$(transmitID, ii, 2)
  TmpMsg.idBytes(j) = Val(subStr)
  j = j + 1
Next ii

If (strLen Mod 2) <> 0 Then
  subStr = "&h" & Mid$(transmitID, 1, 1)
  TmpMsg.idBytes(j) = Val(subStr)
End If

For aa = 0 To 7
  TmpMsg.data(aa) = Val("&h" & Me.data(aa).Text)
Next
TmpMsg.dlc = Val("&h" & Me.dlc.Text)
If TmpMsg.dlc > 8 Then
  TmpMsg.dlc = 8
End If
TmpMsg.flags = 0

ev = Build_vbEvent_tagData_vbMsg(ev, TmpMsg)
If Me.moreInfo.Value Then
  Me.Output.AddItem ">>> Transmit a message"
End If

vErr = vbTransmit(gPortHandle, chanMask, ev)
If vErr Then Fehler

' for checking if a message was transmited
transmited = True
transmitCounter = 10 ' wait 100ms for answer

End Sub

Private Sub Command10_Click()
' FlushTransmitQueue

vErr = vbFlushTransmitQueue(gPortHandle, chanMask)
If vErr Then
   Fehler
Else
   If Me.moreInfo.Value Then
     Me.Output.AddItem ">>> FlushTransmitQueue"
   End If
End If

End Sub

Private Sub Command12_Click()
' Rx using SetNotification

Dim evstr As String * 255

rxc = 1
vErr = vbSetNotification(gPortHandle, h, rxc)
If vErr = VSUCCESS Then
   Me.Output.AddItem ">>> SetNotification, Waiting 3 seconds ..."

   If WaitForSingleObject(h, 3000) <> WAIT_TIMEOUT Then
     Me.Output.AddItem ">>> Event triggered..."
     While schlaufe = 0
     
       vErr = vbReceive1(gPortHandle, pEvent)
       If vErr = VSUCCESS Then
          nix = vbGetEventString(pEvent, evstr)
          Me.Output.AddItem evstr
       ElseIf vErr <> VERR_QUEUE_IS_EMPTY Then
          Fehler
       Else
         Exit Sub
       End If
     Wend
   Else
     Me.Output.AddItem ">>>No Event triggered..."
   End If
Else
   Fehler
End If

End Sub

Public Sub Command13_Click()
vErr = vbClosePort(gPortHandle)
gPortHandle = INVALID_PORTHANDLE
vErr = vbCloseDriver
If Me.moreInfo.Value Then
  Me.Output.AddItem ">>> Close Port"
  Me.Output.AddItem ">>> Close Driver"
End If

End Sub

Private Sub Command14_Click()

messageCount = 0
overrunCount = 0

Me.Counter_Click

End Sub

Private Sub Command16_Click()
' GetReceiveQueueLevel
Dim n As Long
Me.Output.AddItem ">>> Get the receive queue level"

vErr = vbGetReceiveQueueLevel(gPortHandle, n)
If vErr Then
   Fehler
Else
  Me.Output.AddItem ">>> " & n & " event(s) in the receive queue"
End If

End Sub

Private Sub Command2_Click()
' Transmit a message burst

Dim TmpMsg As vbMsg

Dim transmitID As String
transmitID = Me.idSelect.Text
If (Me.extended.Value) Then
  transmitID = Hex$(Val("&h" & Me.idSelect.Text) Or &H80000000)
End If

ev.tag = v_TRANSMIT_MSG
TmpMsg.flags = 0

j = 0
strLen = Len(transmitID)
For ii = strLen - 1 To 1 Step -2
  subStr = "&h" & Mid$(transmitID, ii, 2)
  TmpMsg.idBytes(j) = Val(subStr)
  j = j + 1
Next ii

If (strLen Mod 2) <> 0 Then
  subStr = "&h" & Mid$(transmitID, 1, 1)
  TmpMsg.idBytes(j) = Val(subStr)
End If

For ii = 0 To 7
   TmpMsg.data(ii) = ii
Next
TmpMsg.dlc = 8

ev = Build_vbEvent_tagData_vbMsg(ev, TmpMsg)

'Stop

prompt = "Type duration in msec" & vbCr & vbCr & "Use '#<n>' to generaate n messages"

v = InputBox(prompt, "Sending a message burst")
If InStr(v, "#") Then
   t = Timer
   For i = 1 To Val(Mid$(v, 2))
    vErr = vbTransmit(gPortHandle, chanMask, ev)
    If (vErr And (vErr <> VERR_QUEUE_IS_FULL)) Then
       Fehler
       Exit Sub
    End If
   Next
   MsgBox "Messages sent for " & Timer - t & "s"
Else
    endTimer = Timer + Val(v) / 1000
    z = 0
    While Timer < endTimer
       vErr = vbTransmit(gPortHandle, chanMask, ev)
       z = z + 1
       If (vErr And (vErr <> VERR_QUEUE_IS_FULL)) Then
          Fehler
          Exit Sub
       End If
    Wend
    MsgBox z & " messages sent!", vbOKOnly, "Burst finished"
End If
If Me.moreInfo.Value Then
  Me.Output.AddItem ">>> Sending a message burst"
End If

' for checking if a message was transmited
transmited = True
transmitCounter = 10 ' wait 100ms for answer

End Sub
Private Sub Command3_Click()
' Transmit a remote message

Dim TmpMsg As vbMsg

Dim transmitID As String
transmitID = Me.idSelect.Text
If (Me.extended.Value) Then
  transmitID = Hex$(Val("&h" & Me.idSelect.Text) Or &H80000000)
End If

ev.tag = v_TRANSMIT_MSG

' Workaround for Japan
j = 0
strLen = Len(transmitID)
For ii = strLen - 1 To 1 Step -2
  subStr = "&h" & Mid$(transmitID, ii, 2)
  TmpMsg.idBytes(j) = Val(subStr)
  j = j + 1
Next ii

If (strLen Mod 2) <> 0 Then
  subStr = "&h" & Mid$(transmitID, 1, 1)
  TmpMsg.idBytes(j) = Val(subStr)
End If

For ii = 0 To 7
   TmpMsg.data(ii) = ii
Next
TmpMsg.flags = MSGFLAG_REMOTE_FRAME
TmpMsg.dlc = 8

ev = Build_vbEvent_tagData_vbMsg(ev, TmpMsg)
vErr = vbTransmit(gPortHandle, chanMask, ev)
If Me.moreInfo.Value Then
  Me.Output.AddItem ">>> transmit a remote frame"
End If
If vErr Then Fehler

' for checking if a message was transmited
transmited = True
transmitCounter = 10 ' wait 100ms for answer

End Sub

Private Sub Command5_Click()

Me.Output.Clear

End Sub

Private Sub Command4_Click()
' Request chipstate

vErr = vbRequestChipState(gPortHandle, chanMask)
If vErr Then Fehler
If Me.moreInfo.Value Then
  Me.Output.AddItem ">>> Request Chip State"
End If


End Sub

Private Sub Command6_Click()
' Reset clock

vErr = vbResetClock(gPortHandle)
If vErr Then
   Fehler
Else
  If Me.moreInfo.Value Then
    Me.Output.AddItem ">>> Clock reset"
  End If
End If

End Sub
Private Sub Command8_Click()
' Toggle output mode

If gOutputMode = OUTPUT_MODE_SILENT Then
   gOutputMode = OUTPUT_MODE_NORMAL
   StatusMsg = ">>> Output mode set to normal"
Else
   gOutputMode = OUTPUT_MODE_SILENT
   StatusMsg = ">>> Output mode set to silent"
End If

vErr = vbSetChannelOutput(gPortHandle, chanMask, gOutputMode)
If vErr Then
   Fehler
Else
  Me.Output.AddItem StatusMsg
End If


End Sub

Private Sub Command9_Click()
' Toggle timer

If gTimerRate = 0 Then
   gTimerRate = Val(InputBox("Timer Resolution (10µs)", "SetTimerRate"))
Else
   gTimerRate = 0
End If

vErr = vbSetTimerRate(gPortHandle, "&h" & Hex$(gTimerRate))
If Me.moreInfo.Value Then
  Me.Output.AddItem ">>> Set timer rate"
End If

If vErr Then Fehler


End Sub

Public Sub Counter_Click()

Me.Counter.Caption = "MessageCount " & messageCount & vbCr & vbCr & "OverrunCount " & overrunCount

End Sub

Public Sub DeActivate_Click()
' De-/Activate channel

If Activated Then
   vErr = vbDeactivateChannel(gPortHandle, gChannelMask)
   If vErr Then
      Fehler
   Else
      If Me.moreInfo.Value Then
        Me.Output.AddItem ">>> Channel deactivated"
      End If
      Me.DeActivate.Caption = "Channel activate"
      Activated = False
      
      If gPermissionMask Then
        For tmpChan = 0 To TmpCfg.channelCount - 1
          If resetAcceptance(tmpChan) Then
            Me.acceptance_add(tmpChan).Enabled = True
          End If
          Me.Bitrate_Ch(tmpChan).Enabled = False
          Me.acceptance_from(tmpChan).Enabled = True
          Me.acceptance_region(tmpChan).Enabled = True
          Me.acceptance_reset(tmpChan).Enabled = True
          Me.acceptance_to(tmpChan).Enabled = True
          If (gPermissionMask And (2 ^ tmpChan)) Then
            Me.Bitrate_Ch(tmpChan).Enabled = True
          End If
        Next
        Me.requestTransmit.Enabled = True
      End If
      
   End If
Else
   vErr = vbActivateChannel(gPortHandle, gChannelMask)
   If vErr Then
      Fehler
   Else
      If Me.moreInfo.Value Then
        Me.Output.AddItem ">>> Channel activated " & gChannelMask
      End If
      Me.DeActivate.Caption = "Channel deactivate"
      
      Activated = True
           
      For aa = 0 To 7
        Me.Bitrate_Ch(aa).Enabled = False
        Me.acceptance_add(aa).Enabled = False
        Me.acceptance_from(aa).Enabled = False
        Me.acceptance_reset(aa).Enabled = False
        Me.acceptance_to(aa).Enabled = False
      Next
      
      If gPermissionMask Then
      
        Me.requestTransmit.Enabled = False
      
        For tmpChan = 0 To TmpCfg.channelCount - 1
          If (gPermissionMask And (2 ^ tmpChan)) Then
            vErr = vbSetChannelBitrate(gPortHandle, (gPermissionMask And (2 ^ tmpChan)), Me.Bitrate_Ch(tmpChan).Text)
            If Me.moreInfo.Value Then
              Me.Output.AddItem ">>> Set Channel Bitrate: " & Me.Bitrate_Ch(tmpChan).Text
            End If
          End If
          If vErr Then
            Fehler
          End If
        Next
      End If
      
   End If
End If
CfgUpdate_Click
End Sub
Private Sub Form_Load()

Declarations

h = CreateEvent(vbNullString, False, False, vbNullString)

Me.OnOffline_Click

End Sub
Private Sub Form_Unload(Cancel As Integer)

Me.Command13_Click

End Sub

Public Sub OnOffline_Click()

If Me.OnOffline.Caption = "Close&Driver" Then
   Me.OnOffline.Caption = "Open&Driver"
   Me.Timer1.Enabled = False
   vbClosePort (gPortHandle)
   gPortHandle = INVALID_PORTHANDLE
   vbCloseDriver
   If Me.moreInfo.Value Then
     Me.Output.AddItem ">>> Close Port"
     Me.Output.AddItem ">>> Close Driver"
   End If
Else
   Me.OnOffline.Caption = "Close&Driver"
   vErr = InitDriver
   If vErr Then Fehler
End If

End Sub

Private Sub requestTransmit_Click()
  If Me.requestTransmit.Value Then
    ' generate TX
    vErr = vbSetChannelMode(gPortHandle, gPermissionMask, 1, 0)
  Else
    ' don't generate TX
    vErr = vbSetChannelMode(gPortHandle, gPermissionMask, 0, 0)
  End If
End Sub

Private Sub Timer1_Timer()

If Not Activated Then
   Exit Sub
End If

Dim evstr As String * 255
Dim Tmp As Long
Dim timestamp As Double

If Me.ReceiveMode(0).Value Then
' ////////////////
' // Receive1

'Stop
vErr = vbReceive1(gPortHandle, pEvent)

' check if a message was transmited
If (transmited) Then
  transmitCounter = transmitCounter - 1
  If ((transmitCounter = 0) And Me.requestTransmit.Value) Then
    Me.Output.AddItem ">>> couldn't transmit the message! "
    transmited = False
  End If
End If

If vErr = VSUCCESS Then
   vErr = vbGetEventString(pEvent, evstr)
   Me.Output.AddItem evstr
   
   ' check if a message was transmited
   If (transmited) Then
     If (pEvent.tagData(4) And MSGFLAG_TX) Then
       Me.Output.AddItem ">>> the message was transmited "
       transmited = False
     End If
   End If
   
   messageCount = messageCount + 1
   If pEvent.timestamp Then
      timestamp = pEvent.timestamp
      If (timestamp < 0) Then
        timestamp = timestamp + 4294967296#    'fix signed reperesentation of VB as unsigned
      End If
      If lastTime > timestamp Then
        Me.Output.AddItem "!!! Time decreasing !!!  DeltaT = -" & lastTime - timestamp
      End If
      lastTime = timestamp
   End If
   If Get_vbEvent_tagData_vbMsg(pEvent).flags And MSGFLAG_OVERRUN Then
      overrunCount = overrunCount + 1
   End If
ElseIf vErr <> VERR_QUEUE_IS_EMPTY Then
   Fehler
Else
   Exit Sub
End If


      
End If   ' me.receivemode()

Me.Counter_Click

lc = Me.Output.ListCount
Me.Output.ListIndex = lc - 1


End Sub

Private Sub TransmitID_Select_Click(Index As Integer)

Dim transmitID As String
transmitID = Me.idSelect.Text
If (Me.extended.Value) Then
  transmitID = Hex$(Val("&h" & Me.idSelect.Text) Or &H80000000)
End If

Dim z As Double
z = Val("&h" & transmitID)

f = 1
Select Case Index
Case 0          ' -1
   If z = &H80000000 Then
      z = &H7FFFFFFF
      s = 0
   Else
      s = -1
   End If
Case 1
   If z = &H7FFFFFFF Then
      z = &H80000000
      s = 0
   Else
      s = 1
   End If
Case 2
   f = 0
End Select

z = (z + s) * f

Me.idSelect.Text = Hex$(z)

End Sub

Private Sub ttt_Click(Index As Integer)

End Sub
