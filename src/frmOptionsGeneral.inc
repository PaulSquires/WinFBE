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
' frmOptionsGeneral Window procedure
' ========================================================================================
Function frmOptionsGeneral_WndProc( ByVal HWnd   As HWnd, _
                                    ByVal uMsg   As UINT, _
                                    ByVal wParam As WPARAM, _
                                    ByVal lParam As LPARAM _
                                    ) As LRESULT

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmOptionsGeneral_Show
' ========================================================================================
Function frmOptionsGeneral_Show( ByVal hWndParent As HWnd, _
                                 ByVal nCmdShow   As Long = 0 _
                                 ) As Long

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   HWND_FRMOPTIONSGENERAL = pWindow->Create( hWndParent, "", @frmOptionsGeneral_WndProc, 0, 0, 0, 0, _
        WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   ' Height and width of this child form is set in frmOptions_OnNotify when the treeview option is selected.
   
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSGENERAL_CHKMULTIPLEINSTANCES, L(119,"Allow multiple instances"), 0, 0, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSGENERAL_CHKCOMPILEAUTOSAVE, L(121,"Autosave files before compiling"), 0, 21, 235, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSGENERAL_CHKCLOSEFUNCLIST, L(227,"Close Function List on selection"), 0, 42, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSGENERAL_CHKHIDECOMPILE, L(218,"Hide successful compile results message"), 0, 63, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSGENERAL_CHKASKEXIT, L(274,"Ask before exiting the editor"), 0, 84, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSGENERAL_CHKHIDETOOLBAR, L(339,"Hide main toolbar"), 0, 105, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSGENERAL_CHKHIDESTATUSBAR, L(340,"Hide main statusbar"), 0, 126, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   Dim As Long idx = 0
   Dim As HWnd HWnd = HWND_FRMOPTIONSGENERAL
   
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSGENERAL_CHKMULTIPLEINSTANCES),  gConfig.MultipleInstances)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSGENERAL_CHKCLOSEFUNCLIST),      gConfig.CloseFuncList)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSGENERAL_CHKCOMPILEAUTOSAVE),    gConfig.CompileAutoSave)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSGENERAL_CHKHIDECOMPILE),        gConfig.HideCompile)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSGENERAL_CHKASKEXIT),            gConfig.AskExit)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSGENERAL_CHKHIDETOOLBAR),        gConfig.HideToolbar)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSGENERAL_CHKHIDESTATUSBAR),      gConfig.HideStatusbar)

   Function = 0
   
End Function





