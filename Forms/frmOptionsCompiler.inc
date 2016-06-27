

' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmOptionsCompiler
' ========================================================================================
Function frmOptionsCompiler_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Select Case id

      Case IDC_FRMOPTIONSCOMPILER_CMDFBWIN32, IDC_FRMOPTIONSCOMPILER_CMDFBWIN64, _
           IDC_FRMOPTIONSCOMPILER_CMDFBHELPFILE, IDC_FRMOPTIONSCOMPILER_CMDAPIHELPPATH

         If codeNotify = BN_CLICKED Then
            ' Display the Open File Dialog
            Dim pwszName As WString Ptr = AfxIFileOpenDialogW(HWnd, id)
            If pwszName Then
               If id = IDC_FRMOPTIONSCOMPILER_CMDFBWIN32 Then SetWindowTextW( GetDlgItem(HWnd, IDC_FRMOPTIONSCOMPILER_TXTFBWIN32), pwszName )
               If id = IDC_FRMOPTIONSCOMPILER_CMDFBWIN64 Then SetWindowTextW( GetDlgItem(HWnd, IDC_FRMOPTIONSCOMPILER_TXTFBWIN64), pwszName )
               If id = IDC_FRMOPTIONSCOMPILER_CMDFBHELPFILE Then SetWindowTextW( GetDlgItem(HWnd, IDC_FRMOPTIONSCOMPILER_TXTFBHELPFILE), pwszName )
               If id = IDC_FRMOPTIONSCOMPILER_CMDAPIHELPPATH Then SetWindowTextW( GetDlgItem(HWnd, IDC_FRMOPTIONSCOMPILER_TXTWIN32HELPPATH), pwszName )
               CoTaskMemFree pwszName
            End If
            Exit Function
         End If

   End Select

   Function = 0
End Function


' ========================================================================================
' frmOptionsCompiler Window procedure
' ========================================================================================
Function frmOptionsCompiler_WndProc( ByVal HWnd   As HWnd, _
                                     ByVal uMsg   As UINT, _
                                     ByVal wParam As WPARAM, _
                                     ByVal lParam As LPARAM _
                                     ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_COMMAND, frmOptionsCompiler_OnCommand)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProcW(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmOptionsCompiler_Show
' ========================================================================================
Function frmOptionsCompiler_Show( ByVal hWndParent As HWnd, _
                                  ByVal nCmdShow   As Long = 0 _
                                  ) As Long

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI
   
   HWND_FRMOPTIONSCOMPILER = pWindow->Create( hWndParent, "", @frmOptionsCompiler_WndProc, 0, 0, 408, 297, _
        WS_CHILD, WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->SetClientSize(408, 297)
   
   pWindow->AddControl("TEXTBOX", , IDC_FRMOPTIONSCOMPILER_TXTFBWIN32, "", 0, 20, 306, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("BUTTON", , IDC_FRMOPTIONSCOMPILER_CMDFBWIN32, "...", 311, 20, 24, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMOPTIONSCOMPILER_TXTFBWIN64, "", 0, 68, 306, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("BUTTON", , IDC_FRMOPTIONSCOMPILER_CMDFBWIN64, "...", 311, 68, 24, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMOPTIONSCOMPILER_TXTFBSWITCHES, "", 0, 130, 306, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("TEXTBOX", , IDC_FRMOPTIONSCOMPILER_TXTFBHELPFILE, "", 0, 182, 306, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("BUTTON", , IDC_FRMOPTIONSCOMPILER_CMDFBHELPFILE, "...", 312, 182, 24, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMOPTIONSCOMPILER_TXTWIN32HELPPATH, "", 0, 239, 306, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("BUTTON", , IDC_FRMOPTIONSCOMPILER_CMDAPIHELPPATH, "...", 312, 239, 24, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMOPTIONSCOMPILER_LBLSWITCHES, L(153,"Additional compiler option switches (optional):"), 0, 112, 324, 14, _
        WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMOPTIONSCOMPILER_LBLFBC32, L(154,"FBC 32bit Compiler:"), 0, 2, 239, 15, _
        WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMOPTIONSCOMPILER_LBLFBHELP, L(155,"FreeBASIC Help file (*.chm):"), 0, 164, 304, 14, _
        WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMOPTIONSCOMPILER_LBLFBC64, L(156,"FBC 64bit Compiler:"), 0, 50, 229, 15, _
        WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMOPTIONSCOMPILER_LBLAPIHELP, L(157,"WinAPI Help file (Win32_fb.chm):"), 0, 219, 266, 16, _
        WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   ProcessFromCurdrive(@gConfig.FBWINCompiler32)
   ProcessFromCurdrive(@gConfig.FBWINCompiler64)
   ProcessFromCurdrive(@gConfig.CompilerHelpfile)
   ProcessFromCurdrive(@gConfig.Win32APIHelpfile)

   Dim As HWnd HWnd = HWND_FRMOPTIONSCOMPILER
   SetWindowTextW GetDlgItem( HWnd, IDC_FRMOPTIONSCOMPILER_TXTFBWIN32),       gConfig.FBWINCompiler32
   SetWindowTextW GetDlgItem( HWnd, IDC_FRMOPTIONSCOMPILER_TXTFBWIN64),       gConfig.FBWINCompiler64
   SetWindowTextW GetDlgItem( HWnd, IDC_FRMOPTIONSCOMPILER_TXTFBSWITCHES),    gConfig.CompilerSwitches
   SetWindowTextW GetDlgItem( HWnd, IDC_FRMOPTIONSCOMPILER_TXTFBHELPFILE),    gConfig.CompilerHelpfile
   SetWindowTextW GetDlgItem( HWnd, IDC_FRMOPTIONSCOMPILER_TXTWIN32HELPPATH), gConfig.Win32APIHelpfile

   Function = 0 
End Function
