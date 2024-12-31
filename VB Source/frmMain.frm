VERSION 5.00
Object = "{22BE512E-E6B6-11D2-9BB5-00A0CC3AD9E7}#1.0#0"; "PVOutlookBar.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{85202277-6C76-4228-BC56-7B3E69E8D5CA}#5.0#0"; "IGToolBars50.ocx"
Object = "{1416D7C5-8A28-11CF-9236-444553540000}#8.0#0"; "PVXPLORE8.ocx"
Begin VB.Form frmMain 
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   14670
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   14670
   StartUpPosition =   3  'Windows Default
   Begin PVExplorerLib.PVExplorer MainExplorer 
      Height          =   435
      Left            =   690
      TabIndex        =   2
      Top             =   9630
      Visible         =   0   'False
      Width           =   255
      _Version        =   524288
      LabelEdit       =   0   'False
      Indentation     =   0
      SourceChannel1  =   ""
      TargetChannel1  =   ""
      PathSeparator   =   ""
      Image1          =   "frmMain.frx":0000
      SingleExpand    =   -1  'True
      SourceChannel2  =   ""
      TargetChannel2  =   ""
      Image2          =   "frmMain.frx":037E
      Image3          =   "frmMain.frx":227C
      FileName        =   ""
      LeftPaneWidth   =   240
      DataMember      =   ""
      DataField0      =   ""
      DataField1      =   ""
      DataField2      =   ""
      DataField3      =   ""
      DataField4      =   ""
      DataField5      =   ""
      DataField6      =   ""
      DataField7      =   ""
      DataField8      =   ""
      DataField9      =   ""
      DataField10     =   ""
      DataField11     =   ""
      DataField12     =   ""
      DataField13     =   ""
      DataField14     =   ""
      DataField15     =   ""
      DataField16     =   ""
      DataField17     =   ""
      DataField18     =   ""
      DataField19     =   ""
      PaneDisplay     =   1
      CaptionMode     =   1
      Appearance      =   1
      _ExtentX        =   450
      _ExtentY        =   767
      _StockProps     =   70
   End
   Begin VB.Timer MainTimer 
      Interval        =   500
      Left            =   960
      Top             =   9660
   End
   Begin MSComctlLib.ImageList MainImageList 
      Left            =   90
      Top             =   9630
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":32F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3742
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B94
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3FE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4438
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":488A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4CDC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar MainStatusBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   9870
      Width           =   14670
      _ExtentX        =   25876
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Idle."
            TextSave        =   "Idle."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "DSS Offline"
            TextSave        =   "DSS Offline"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "GPS Offline"
            TextSave        =   "GPS Offline"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "NeoVI Offline"
            TextSave        =   "NeoVI Offline"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Database Offline"
            TextSave        =   "Database Offline"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Log File Offline"
            TextSave        =   "Log File Offline"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "4/12/2004"
         EndProperty
      EndProperty
   End
   Begin ActiveResizer.SSResizer MainResizer 
      Left            =   1050
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   8
      MaxFontSize     =   100
      DesignWidth     =   14670
      DesignHeight    =   10215
   End
   Begin ActiveToolBars.SSActiveToolBars MainToolBar 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      ToolBarsCount   =   2
      ToolsCount      =   4
      PersonalizedMenus=   0
      ActiveColors    =   -1  'True
      DisplayContextMenu=   0   'False
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmMain.frx":4FF6
      ToolBars        =   "frmMain.frx":76AC
   End
   Begin OUTLOOKBARLibCtl.PVOutlookBar OutlookBarMain 
      Height          =   9825
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   14595
      _Version        =   131073
      SoundEffects    =   -1  'True
      Appearance      =   1
      BorderWidth     =   1
      SplitterWindow  =   -1  'True
      SplitterWidth   =   6
      GroupPopupMenu  =   0   'False
      RenameGroups    =   0   'False
      AddGroups       =   0   'False
      RenameItems     =   0   'False
      RemoveGroups    =   0   'False
      AddItems        =   0   'False
      RemoveItems     =   0   'False
      HideOutlookBar  =   0   'False
      ItemPopupMenu   =   -1  'True
      OpenItem        =   -1  'True
      Properties      =   0   'False
      SizeIcons       =   0   'False
      BackColor       =   -2147483636
      TextColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DyanmicResize   =   -1  'True
      UseChildWindows =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()

    
    '---- Configure the Data Explorer
    Dim RootNode As pvxNode, RootNode2 As pvxNode
    Dim Node As pvxNode
    With MainExplorer.Nodes
        
        Set RootNode = .AddRootNode("Data Acq.", 0, 1)
        Set Node = .AddChild(RootNode, "Data", 0, 1)
        Node.Key = "Data"
        
        Set Node = .AddChild(RootNode, "Graphs", 0, 1)
        Node.Key = "Graphs"
        
        Set RootNode2 = .AddRootNode("Preferences", 0, 1)
        
        Set Node = .AddChild(RootNode2, "DSS", 0, 1)
        Node.WindowObject = frmDSSSetup.hWnd
        Node.Key = "DSS"
    
        Set Node = .AddChild(RootNode2, "GPS", 0, 1)
        Node.WindowObject = frmGPSSetup.hWnd
        Node.Key = "GPS"
    
        Set Node = .AddChild(RootNode2, "NeoVI", 0, 1)
        Node.WindowObject = frmNeoVISetup.hWnd
        Node.Key = "NeoVI"
        
        Set Node = .AddChild(RootNode2, "CANCardXL", 0, 1)
        Node.WindowObject = frmVectorCANXL.hWnd
        Node.Key = "Vector"
        
    End With
    
    MainExplorer.SelectedNode = RootNode
    
    
    '---- Configure Outlook BAR
    Dim Group As PVOutlookGroup
    Dim Item As PVOutlookItem
    OutlookBarMain.LargeImageList = MainImageList
    
    
    Set Group = OutlookBarMain.Groups.Add("Data Acq.")
    Set Item = Group.Items.Add("Data", 5)
    Item.ActiveXObject1 = MainExplorer.object
    Item.ExplorerNode = MainExplorer.Nodes.NodeFromKey("Data")
    Item.Display
    
    Set Item = Group.Items.Add("Graphs", 8)
    Item.ActiveXObject1 = MainExplorer.object
    Item.ExplorerNode = MainExplorer.Nodes.NodeFromKey("Graphs")
    
    Set Group = OutlookBarMain.Groups.Add("Preferences")
    
    Set Item = Group.Items.Add("DSS", 0)
    Item.ActiveXObject1 = MainExplorer.object
    Item.ExplorerNode = MainExplorer.Nodes.NodeFromKey("DSS")
    
    
    Set Item = Group.Items.Add("GPS", 2)
    Item.ActiveXObject1 = MainExplorer.object
    Item.ExplorerNode = MainExplorer.Nodes.NodeFromKey("GPS")
    
    Set Item = Group.Items.Add("NeoVI", 1)
    Item.ActiveXObject1 = MainExplorer.object
    Item.ExplorerNode = MainExplorer.Nodes.NodeFromKey("NeoVI")
        
    Set Item = Group.Items.Add("CANCardXL", 1)
    Item.ActiveXObject1 = MainExplorer.object
    Item.ExplorerNode = MainExplorer.Nodes.NodeFromKey("Vector")
    
    
    Me.Caption = "ToVAAS v" & App.Major & "." & App.Minor & "." & App.Revision
    
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If MsgBox("Are you sure?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "Exit") = vbYes Then
        UninstallExceptionHandler
        Unload frmDSSSetup
        Unload frmGPSSetup
        Unload frmNeoVISetup
        Unload frmOverLoads
        Unload frmVectorCANXL
        Unload Me
    Else
        Cancel = 1
    End If
    
End Sub

Private Sub MainTimer_Timer()
    '---- Update the Status Display
    With MainStatusBar
        .Panels(1).Text = "Idle."
        '----- Temporary Override .Panels(2).Text = IIf(frmDSSSetup.IsConnected, "DSS Online", "DSS Offline")
        .Panels(3).Text = IIf(frmGPSSetup.IsConnected, "GPS Online", "GPS Offline")
        .Panels(4).Text = IIf(frmNeoVISetup.IsConnected, "NeoVI Online", "NeoVI Offline")
        .Panels(5).Text = IIf(DB.IsConnected, "Database Online", "Database Offline")
        .Panels(6).Text = IIf(IsLogFileConnected, "LogFile Online", "LogFile Offline")
    End With
End Sub

Private Sub MainToolBar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "ID_New"
        Case "ID_Open"
        
        Case "ID_Exit"
            Unload Me
    End Select
End Sub

Private Sub OutlookBarMain_ItemClick(ByVal Group As OUTLOOKBARLibCtl.IPVOutlookGroup, ByVal Item As OUTLOOKBARLibCtl.IPVOutlookItem)

'    MsgBox "The Item " + Item.Text + " in Group " + Group.Text + " has been clicked"


End Sub
