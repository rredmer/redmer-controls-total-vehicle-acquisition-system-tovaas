Attribute VB_Name = "MainModule"
'****************************************************************************
'**                                                                        **
'** Project....: MTU                                                       **
'**                                                                        **
'** Module.....: MainModule                                                **
'**                                                                        **
'** Description: Initializes global variables & forms to start the app.    **
'**                                                                        **
'** History....:                                                           **
'**    03/20/03 v1.71 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2003 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit

Global DB As DataBaseInterface

Public Sub Main()

    InstallExceptionHandler
    
    InitLogFile
    
    Set DB = New DataBaseInterface
    If DB.OpenDatabase(App.Path & "\Settings.mdb") = True Then
        Load frmDSSSetup
        Load frmGPSSetup
        Load frmNeoVISetup
        Load frmVectorCANXL
        Load frmMain
        frmMain.Show
    End If
    
End Sub
