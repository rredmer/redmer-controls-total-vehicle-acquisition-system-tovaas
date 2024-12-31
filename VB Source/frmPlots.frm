VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Begin VB.Form frmPlots 
   Caption         =   "Plots"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   8910
   Begin ChartfxLibCtl.ChartFX ChartFX1 
      Height          =   2475
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4215
      _cx             =   7435
      _cy             =   4366
      Build           =   20
      TypeMask        =   109576193
      MarkerShape     =   0
      AxesStyle       =   3
      Axis(0).TickMark=   -32767
      Axis(0).GridColor=   16777215
      Axis(2).Min     =   0
      Axis(2).Max     =   100
      Axis(2).GridColor=   16777215
      RGBBk           =   0
      nColors         =   16
      Pallete         =   "frmPlots.frx":0000
      Colors          =   "frmPlots.frx":00E4
      Axis(2).FontColor=   16777215
      Axis(0).FontColor=   16777215
      BorderS         =   8
   End
   Begin ChartfxLibCtl.ChartFX ChartFX1 
      Height          =   2475
      Index           =   1
      Left            =   4320
      TabIndex        =   1
      Top             =   60
      Width           =   4215
      _cx             =   7435
      _cy             =   4366
      Build           =   20
      TypeMask        =   109576193
      MarkerShape     =   0
      AxesStyle       =   3
      Axis(0).TickMark=   -32767
      Axis(0).GridColor=   16777215
      Axis(2).Min     =   0
      Axis(2).Max     =   100
      Axis(2).GridColor=   16777215
      RGBBk           =   0
      nColors         =   16
      Pallete         =   "frmPlots.frx":0184
      Colors          =   "frmPlots.frx":0268
      Axis(2).FontColor=   16777215
      Axis(0).FontColor=   16777215
      BorderS         =   8
   End
   Begin ChartfxLibCtl.ChartFX ChartFX1 
      Height          =   2475
      Index           =   2
      Left            =   60
      TabIndex        =   2
      Top             =   2520
      Width           =   4215
      _cx             =   7435
      _cy             =   4366
      Build           =   20
      TypeMask        =   109576193
      MarkerShape     =   0
      AxesStyle       =   3
      Axis(0).TickMark=   -32767
      Axis(0).GridColor=   16777215
      Axis(2).Min     =   0
      Axis(2).Max     =   100
      Axis(2).GridColor=   16777215
      RGBBk           =   0
      nColors         =   16
      Pallete         =   "frmPlots.frx":0308
      Colors          =   "frmPlots.frx":03EC
      Axis(2).FontColor=   16777215
      Axis(0).FontColor=   16777215
      BorderS         =   8
   End
   Begin ChartfxLibCtl.ChartFX ChartFX1 
      Height          =   2475
      Index           =   3
      Left            =   4320
      TabIndex        =   3
      Top             =   2520
      Width           =   4215
      _cx             =   7435
      _cy             =   4366
      Build           =   20
      TypeMask        =   109576193
      MarkerShape     =   0
      AxesStyle       =   3
      Axis(0).TickMark=   -32767
      Axis(0).GridColor=   16777215
      Axis(2).Min     =   0
      Axis(2).Max     =   100
      Axis(2).GridColor=   16777215
      RGBBk           =   0
      nColors         =   16
      Pallete         =   "frmPlots.frx":048C
      Colors          =   "frmPlots.frx":0570
      Axis(2).FontColor=   16777215
      Axis(0).FontColor=   16777215
      BorderS         =   8
      _Data_          =   "frmPlots.frx":0610
   End
End
Attribute VB_Name = "frmPlots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()

    ChartFX1(0).Top = 0
    ChartFX1(0).Left = 0
    ChartFX1(0).Width = Me.ScaleWidth / 2
    ChartFX1(0).Height = Me.ScaleHeight / 2
    
    ChartFX1(1).Top = 0
    ChartFX1(1).Left = Me.ScaleWidth / 2
    ChartFX1(1).Width = Me.ScaleWidth / 2
    ChartFX1(1).Height = Me.ScaleHeight / 2
    
    ChartFX1(2).Top = Me.ScaleHeight / 2
    ChartFX1(2).Left = 0
    ChartFX1(2).Width = Me.ScaleWidth / 2
    ChartFX1(2).Height = Me.ScaleHeight / 2
    
    ChartFX1(3).Top = Me.ScaleHeight / 2
    ChartFX1(3).Left = Me.ScaleWidth / 2
    ChartFX1(3).Width = Me.ScaleWidth / 2
    ChartFX1(3).Height = Me.ScaleHeight / 2

End Sub
