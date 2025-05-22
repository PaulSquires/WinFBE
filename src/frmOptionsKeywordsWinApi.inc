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

#include once "frmOptionsKeywordsWinApi.bi"


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmOptionsKeywordsWinApi
' ========================================================================================
Function frmOptionsKeywordsWinApi_OnCommand( _
            ByVal HWnd As HWnd, _
            ByVal id As Long, _
            ByVal hwndCtl As HWnd, _
            ByVal codeNotify As UINT _
            ) As LRESULT

   Select Case id
      Case IDC_FRMOPTIONSKEYWORDSWINAPI_TXTKEYWORDS
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
' Process WM_DESTROY message for window/dialog: frmOptionsKeywordsWinApi
' ========================================================================================
Function frmOptionsKeywordsWinApi_OnDestroy( byval HWnd As HWnd ) As LRESULT
   Dim As HFONT hFont = AfxGetWindowFont(GetDlgItem(HWnd, IDC_FRMOPTIONSKEYWORDSWINAPI_TXTKEYWORDS))
   DeleteFont(hFont)
   Function = 0
End Function


' ========================================================================================
' frmOptionsKeywordsWinApi Window procedure
' ========================================================================================
Function frmOptionsKeywordsWinApi_WndProc( _
            ByVal HWnd   As HWnd, _
            ByVal uMsg   As UINT, _
            ByVal wParam As WPARAM, _
            ByVal lParam As LPARAM _
            ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_COMMAND, frmOptionsKeywordsWinApi_OnCommand)
      HANDLE_MSG (HWnd, WM_DESTROY, frmOptionsKeywordsWinApi_OnDestroy)
   End Select

   Function = DefWindowProc( HWnd, uMsg, wParam, lParam )

End Function


' ========================================================================================
' frmOptionsKeywordsWinApi_Show
' ========================================================================================
function frmOptionsKeywordsWinApi_Show( ByVal hWndParent As HWnd ) as LRESULT

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   HWND_FRMOPTIONSKEYWORDSWINAPI = pWindow->Create( hWndParent, "", _
        @frmOptionsKeywordsWinApi_WndProc, 0, 0, 0, 0, _
        WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING )
   ' Height and width of this child form is set in frmOptions_OnNotify when the treeview option is selected.
   
   Dim As HWnd hTextBox = _
   pWindow->AddControl("TEXTBOX", , IDC_FRMOPTIONSKEYWORDSWINAPI_TXTKEYWORDS, "", 0, 0, 386, 345, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or ES_LEFT Or _
        ES_MULTILINE Or ES_AUTOVSCROLL Or ES_WANTRETURN, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING )
   SendMessage( hTextBox, EM_SETLIMITTEXT, (1024 * 100), 0 )    ' max 100K text   
   Dim As HFONT hFont = pWindow->CreateFont("Courier New", 9)
   AfxSetWindowFont( hTextBox, hFont, True )

   gConfig.bKeywordsDirty = False
   AfxSetWindowText( hTextBox, WStr(gConfig.WinApiKeywords) )
   
   Function = 0
   
End Function

