'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2018 Paul Squires, PlanetSquires Software
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
' Process WM_CREATE message for window/dialog: frmCommandLine
' ========================================================================================
Function frmCommandLine_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN

   ' Retrieve a reference to the CWindow class from the CREATESTRUCT structure
   'Dim pCreateStruct As CREATESTRUCT Ptr = lpCreateStructPtr
   'Dim pWindow As CWindow Ptr = Cast(CWindow Ptr, pCreateStruct->lpCreateParams)

   ' This is a modal popup window so disable the parent window
   EnableWindow GetParent(HWnd), False

   '  Message cracker macro expects a True to be returned for a successful
   '  OnCreate handler even though returning -1 from a standard WM_CREATE
   '  call would stop creating the window. This is just one of those Windows
   '  inconsistencies.
   Return True
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmCommandLine
' ========================================================================================
Function frmCommandLine_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Select Case id
      Case IDOK
         If codeNotify = BN_CLICKED Then
            gApp.Projects(gApp.GetActiveProjectIndex).ProjectCommandLine = AfxGetWindowText( GetDlgItem(HWnd, IDC_FRMCOMMANDLINE_TEXTBOX1) )
            SendMessageW HWnd, WM_CLOSE, 0, 0
            Exit Function
         End If
      Case IDCANCEL
         If codeNotify = BN_CLICKED Then
            SendMessageW HWnd, WM_CLOSE, 0, 0
            Exit Function
         End If
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmCommandLine
' ========================================================================================
Function frmCommandLine_OnClose(HWnd As HWnd) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableWindow GetParent(HWnd), True 
   DestroyWindow HWnd
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmCommandLine
' ========================================================================================
Function frmCommandLine_OnDestroy(HWnd As HWnd) As LRESULT
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' frmCommandLine Window procedure
' ========================================================================================
Function frmCommandLine_WndProc( ByVal HWnd   As HWnd, _
                             ByVal uMsg   As UINT, _
                             ByVal wParam As WPARAM, _
                             ByVal lParam As LPARAM _
                             ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,   frmCommandLine_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,    frmCommandLine_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmCommandLine_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,  frmCommandLine_OnCommand)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProcW(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmCommandLine_Show
' ========================================================================================
Function frmCommandLine_Show( ByVal hWndParent As HWnd, _
                              ByVal nCmdShow   As Long = 0 _
                              ) As Long

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   pWindow->Create( hWndParent, L(178,"Command Line"), @frmCommandLine_WndProc, 0, 0, 0, 0, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT)
   pWindow->SetClientSize(420, 100)
   pWindow->Center(pWindow->hWindow, hWndParent)
      
   pWindow->AddControl("LABEL", , IDC_FRMCOMMANDLINE_LABEL1, L(179,"Enter command line arguments:"), 10, 10, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or SS_LEFT, WS_EX_LEFT Or WS_EX_LTRREADING)

   Dim As HWnd hTextBox = _
   pWindow->AddControl("TEXTBOX", , IDC_FRMCOMMANDLINE_TEXTBOX1, "", 10, 33, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("BUTTON", , IDOK, L(0,"&OK"), 252, 63, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_DEFPUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   
   pWindow->AddControl("BUTTON", , IDCANCEL, L(1,"&Cancel"), 336, 63, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   SetWindowText( hTextBox, gApp.Projects(gApp.GetActiveProjectIndex).ProjectCommandLine )
   
   SetFocus hTextBox
   
   ' Process Windows messages
   Function = pWindow->DoEvents(nCmdShow)
   
   ' Delete the frmCommandLine CWindow class manually allocated memory 
   Delete pWindow

End Function
