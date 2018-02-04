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
' Process WM_COMMAND message for window/dialog: frmOptionsKeywords
' ========================================================================================
Function frmOptionsKeywords_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Select Case id
      Case IDC_FRMOPTIONSKEYWORDS_TXTKEYWORDS
         If codeNotify = EN_CHANGE Then 
            ' The EN_CHANGE notification code is not sent when the ES_MULTILINE style is used and the text is sent through WM_SETTEXT. 
            ' Set the flag in gConfig to indicate that the keywords need to be saved to disk should
            ' the user press OK to close the Environment Options dialog.
            gConfig.bKeywordsDirty = True
         End If
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmOptionsKeywords
' ========================================================================================
Function frmOptionsKeywords_OnDestroy(HWnd As HWnd) As LRESULT
   Dim As HFONT hFont = AfxGetWindowFont(GetDlgItem(HWnd, IDC_FRMOPTIONSKEYWORDS_TXTKEYWORDS))
   DeleteFont(hFont)
   Function = 0
End Function


' ========================================================================================
' frmOptionsKeywords Window procedure
' ========================================================================================
Function frmOptionsKeywords_WndProc( ByVal HWnd   As HWnd, _
                                  ByVal uMsg   As UINT, _
                                  ByVal wParam As WPARAM, _
                                  ByVal lParam As LPARAM _
                                  ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_COMMAND, frmOptionsKeywords_OnCommand)
      HANDLE_MSG (HWnd, WM_DESTROY, frmOptionsKeywords_OnDestroy)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProcW(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmOptionsKeywords_Show
' ========================================================================================
Function frmOptionsKeywords_Show( ByVal hWndParent As HWnd, _
                                  ByVal nCmdShow   As Long = 0 _
                                  ) As Long

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   HWND_FRMOPTIONSKEYWORDS = pWindow->Create( hWndParent, "", @frmOptionsKeywords_WndProc, 0, 0, 0, 0, _
        WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   ' Height and width of this child form is set in frmOptions_OnNotify when the treeview option is selected.
   
   Dim As HWnd hTextBox = _
   pWindow->AddControl("TEXTBOX", , IDC_FRMOPTIONSKEYWORDS_TXTKEYWORDS, "", 0, 0, 386, 345, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or ES_LEFT Or ES_MULTILINE Or ES_AUTOVSCROLL Or ES_WANTRETURN, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   
   Dim As HFONT hFont = pWindow->CreateFont("Courier New", 9)
   AfxSetWindowFont hTextBox, hFont, True

   gConfig.bKeywordsDirty = False
   SetWindowTextW hTextBox, WStr(gConfig.FBKeywords)
   
   Function = 0
   
End Function

