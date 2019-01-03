'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2019 Paul Squires, PlanetSquires Software
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT any WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS for A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmAbout
' ========================================================================================
Function frmAbout_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN

   ' This is a modal popup window so disable the parent window
   DisableAllModeless()

   '  Message cracker macro expects a True to be returned for a successful
   '  OnCreate handler even though returning -1 from a standard WM_CREATE
   '  call would stop creating the window. This is just one of those Windows
   '  inconsistencies.
   Return True
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmAbout
' ========================================================================================
Function frmAbout_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Select Case id
      Case IDC_FRMABOUT_CMDUPDATES
         If codeNotify = BN_CLICKED Then
            DoCheckForUpdates
         End If
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmAbout
' ========================================================================================
Function frmAbout_OnClose(HWnd As HWnd) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableAllModeless()
   DestroyWindow HWnd
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmAbout
' ========================================================================================
Function frmAbout_OnDestroy(HWnd As HWnd) As LRESULT
   Dim As HFONT hFont = AfxGetWindowFont(GetDlgItem(HWnd, IDC_FRMABOUT_LBLAPPNAME))
   DeleteFont(hFont)
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' frmAbout Window procedure
' ========================================================================================
Function frmAbout_WndProc( ByVal HWnd   As HWnd, _
                           ByVal uMsg   As UINT, _
                           ByVal wParam As WPARAM, _
                           ByVal lParam As LPARAM _
                           ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,  frmAbout_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,   frmAbout_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY, frmAbout_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND, frmAbout_OnCommand)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmAbout_Show
' ========================================================================================
Function frmAbout_Show( ByVal hWndParent As HWnd, _
                        ByVal nCmdShow   As Long = 0 _
                        ) As Long

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   ' Add an image control                                                         
   DIM pImageCtx AS CImageCtx = CImageCtx(pWindow, IDC_FRMIMAGEMANAGER_IMAGECTX, , 30, 30, 100, 100)
   pImageCtx.LoadImageFromResource( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON" )


   pWindow->Create( hWndParent, L(74,"About"), @frmAbout_WndProc, 0, 0, 0, 0, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT )
   pWindow->SetClientSize(625, 460)
   pWindow->Center(pWindow->hWindow, hWndParent)
      
   Dim As HWnd hLabel = _
      pWindow->AddControl("LABEL", , IDC_FRMABOUT_LBLAPPNAME, APPNAME, 210, 30, 400, 30, _
           WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_CENTERIMAGE Or SS_NOTIFY, _
           WS_EX_LEFT Or WS_EX_LTRREADING)
   Dim As HFONT hFont = pWindow->CreateFont("", 12, FW_BOLD)
   AfxSetWindowFont hLabel, hFont, True

   pWindow->AddControl("LABEL", , IDC_FRMABOUT_LBLAPPVERSION, _
        "Version " & APPVERSION & APPBITS, 210, 75, 200, 20, _
        WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("LABEL", , IDC_FRMABOUT_LBLAPPCOPYRIGHT, _
        APPCOPYRIGHT, 210, 90, 300, 20, _
        WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("BUTTON", , IDC_FRMABOUT_CMDUPDATES, L(91,"Check for Updates"), 210, 140, 200, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
    
'                       "Special thanks to:" & vbcrlf & _
'                       "José Roca (WinFBX Framework)", _

   ' Process Windows messages
   Function = pWindow->DoEvents(SW_SHOW)
   
   ' Delete the frmAbout CWindow class manually allocated memory 
   Delete pWindow

End Function


