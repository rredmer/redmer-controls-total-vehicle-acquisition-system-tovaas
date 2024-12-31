VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmNeoSetup 
   Caption         =   "NeoVI Setup"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6690
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
   ScaleHeight     =   5205
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   555
      Left            =   840
      TabIndex        =   3
      Top             =   4620
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   555
      Left            =   3840
      TabIndex        =   2
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   555
      Left            =   5280
      TabIndex        =   1
      Top             =   4560
      Width           =   1335
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   4395
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6495
      _Version        =   393216
      _ExtentX        =   11456
      _ExtentY        =   7752
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmNeoSetup.frx":0000
   End
End
Attribute VB_Name = "frmNeoSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    Unload Me

End Sub

Private Sub cmdTest_Click()

    frmNeoVIExample.Show vbModal
    
End Sub

Private Sub Form_Load()

End Sub
