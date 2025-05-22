'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2025 Paul Squires, PlanetSquires Software
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

#include once "frmAbout.bi"

' ========================================================================================
' Process WM_CREATE message for window/dialog: frmAbout
' ========================================================================================
Function frmAbout_OnCreate( _
            ByVal HWnd As HWnd, _
            ByVal lpCreateStructPtr As LPCREATESTRUCT _
            ) As BOOLEAN

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
Function frmAbout_OnCommand( _
            ByVal HWnd As HWnd, _
            ByVal id As Long, _
            ByVal hwndCtl As HWnd, _
            ByVal codeNotify As UINT _
            ) As LRESULT

   Select Case id
      Case IDC_FRMABOUT_CMDUPDATES
         If codeNotify = BN_CLICKED Then
            DoCheckForUpdates( hwnd, false )
         End If
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmAbout
' ========================================================================================
Function frmAbout_OnClose( byval HWnd As HWnd) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableAllModeless()
   DestroyWindow HWnd
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmAbout
' ========================================================================================
Function frmAbout_OnDestroy( byval HWnd As HWnd) As LRESULT
   Dim As HFONT hFont = AfxGetWindowFont(GetDlgItem(HWnd, IDC_FRMABOUT_LBLAPPNAME))
   DeleteFont(hFont)
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' frmAbout Window procedure
' ========================================================================================
Function frmAbout_WndProc( _
            ByVal HWnd   As HWnd, _
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
Function frmAbout_Show( ByVal hWndParent As HWnd ) as LRESULT

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI


   pWindow->Create( hWndParent, L(74,"About"), @frmAbout_WndProc, 0, 0, 0, 0, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT )
   pWindow->SetClientSize(580, 280)
   pWindow->Center(pWindow->hWindow, hWndParent)
      
   ' Add an image control                                                         
   DIM pImageCtx AS CImageCtx = CImageCtx(pWindow, IDC_FRMABOUT_IMAGE1, , 60, 50, 100, 100)
   pImageCtx.LoadImageFromResource( pWindow->InstanceHandle, "IMAGE_WINFBE" )

   Dim As HWnd hLabel = _
      pWindow->AddControl("LABEL", , IDC_FRMABOUT_LBLAPPNAME, APPNAME, 200, 40, 400, 30, _
           WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_CENTERIMAGE Or SS_NOTIFY, WS_EX_LEFT Or WS_EX_LTRREADING)
   Dim As HFONT hFont = pWindow->CreateFont("", 12, FW_BOLD)
   AfxSetWindowFont hLabel, hFont, True

   pWindow->AddControl("LABEL", , IDC_FRMABOUT_LBLAPPVERSION, "Version " & APPVERSION & APPBITS, 200, 71, 200, 20, _
        WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_NOTIFY, WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("LABEL", , IDC_FRMABOUT_LBLAPPCOPYRIGHT, APPCOPYRIGHT, 200, 115, 370, 20, _
        WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_NOTIFY, WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("LABEL", , IDC_FRMABOUT_LBLJOSE, "Special thanks to: José Roca (WinFBX Framework)", _
        200, 131, 300, 20, WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_NOTIFY, WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("BUTTON", , IDC_FRMABOUT_CMDUPDATES, L(91,"Check for Updates"), 200, 180, 200, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, WS_EX_LEFT Or WS_EX_LTRREADING)
    

   ' Process Windows messages
   Function = pWindow->DoEvents(SW_SHOW)
   
   ' Delete the frmAbout CWindow class manually allocated memory 
   Delete pWindow

End Function


