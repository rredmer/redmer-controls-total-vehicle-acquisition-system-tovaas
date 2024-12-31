VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Begin VB.Form frmOverLoads 
   Caption         =   "Over Loads"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   9300
   Begin ChartfxLibCtl.ChartFX chtOverloads 
      Height          =   1995
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9015
      _cx             =   15901
      _cy             =   3519
      Build           =   20
      TypeMask        =   109576194
      Style           =   -9438215
      Volume          =   80
      AxesStyle       =   0
      nColors         =   16
      Colors          =   "frmOverloads.frx":0000
      nPts            =   20
      nSer            =   1
      NumPoint        =   20
      NumSer          =   1
      _Data_          =   "frmOverloads.frx":00A0
   End
End
Attribute VB_Name = "frmOverLoads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()

    ' Resize the Bar Graph
    chtOverloads.Top = 0
    chtOverloads.Left = 0
    chtOverloads.Height = Me.ScaleHeight
    chtOverloads.Width = Me.ScaleWidth
    
End Sub
