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
' Process WM_COMMAND message for window/dialog: frmOptionsLocal
' ========================================================================================
Function frmOptionsLocal_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Select Case id

      Case IDC_FRMOPTIONSLOCAL_CMDLOCALIZATION

         If codeNotify = BN_CLICKED Then
            ' Display the Open File Dialog
            Dim pwszName As WString Ptr = AfxIFileOpenDialogW(HWnd, id)
            If pwszName Then
               Dim wzFilename As WString * MAX_PATH 
               wzFilename = AfxStrPathname( "NAMEX", *pwszName )
               SetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSLOCAL_LBLLOCALIZATION), wzFilename )
               CoTaskMemFree pwszName
            End If
            Exit Function
         End If

   End Select

   Function = 0
End Function


' ========================================================================================
' frmOptionsLocal Window procedure
' ========================================================================================
Function frmOptionsLocal_WndProc( ByVal HWnd   As HWnd, _
                                  ByVal uMsg   As UINT, _
                                  ByVal wParam As WPARAM, _
                                  ByVal lParam As LPARAM _
                                  ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_COMMAND, frmOptionsLocal_OnCommand)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProcW(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmOptionsLocal_Show
' ========================================================================================
Function frmOptionsLocal_Show( ByVal hWndParent As HWnd, _
                                ByVal nCmdShow   As Long = 0 _
                                ) As Long

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   HWND_FRMOPTIONSLOCAL = pWindow->Create( hWndParent, "", @frmOptionsLocal_WndProc, 0, 0, 0, 0, _
        WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   ' Height and width of this child form is set in frmOptions_OnNotify when the treeview option is selected.
   
   pWindow->AddControl("BUTTON", , IDC_FRMOPTIONSLOCAL_CMDLOCALIZATION, L(126,"Select"), 265, 16, 75, 26, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMOPTIONSLOCAL_LBLLOCALIZATION, "English.lang", 14, 19, 226, 22, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_CENTER Or SS_CENTERIMAGE Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("GROUPBOX", , IDC_FRMOPTIONSLOCAL_FRAMELOCALIZATION, L(129,"Localization"), 0, 0, 385, 50, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)

   SetWindowTextW GetDlgItem( HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_LBLLOCALIZATION), gConfig.LocalizationFile
   
   Function = 0
   
End Function

