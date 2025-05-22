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

#include once "frmOptionsCompiler.bi"
#include once "clsConfig.bi"


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmOptionsCompiler
' ========================================================================================
private Function frmOptionsCompiler_OnCommand( ByVal HWnd As HWnd, _
                                               ByVal id As Long, _
                                               ByVal hwndCtl As HWnd, _
                                               ByVal codeNotify As UINT _
                                               ) As LRESULT

   Select Case id

      case IDC_FRMOPTIONSCOMPILER_CMDFBHELPFILE, IDC_FRMOPTIONSCOMPILER_CMDWINFBXHELPPATH

         If codeNotify = BN_CLICKED Then
            ' Display the Open File Dialog
            Dim pwszName As WString Ptr = AfxIFileOpenDialogW(HWnd, id)
            If pwszName Then
               If id = IDC_FRMOPTIONSCOMPILER_CMDFBHELPFILE Then SetWindowTextW( GetDlgItem(HWnd, IDC_FRMOPTIONSCOMPILER_TXTFBHELPFILE), pwszName )
               If id = IDC_FRMOPTIONSCOMPILER_CMDWINFBXHELPPATH Then SetWindowTextW( GetDlgItem(HWnd, IDC_FRMOPTIONSCOMPILER_TXTWINFBXHELPPATH), pwszName )
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
private Function frmOptionsCompiler_WndProc( ByVal HWnd   As HWnd, _
                                             ByVal uMsg   As UINT, _
                                             ByVal wParam As WPARAM, _
                                             ByVal lParam As LPARAM _
                                             ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_COMMAND, frmOptionsCompiler_OnCommand)
   End Select

   Function = DefWindowProcW(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmOptionsCompiler_Show
' ========================================================================================
public Function frmOptionsCompiler_Show( ByVal hWndParent As HWnd ) as LRESULT

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI
   
   HWND_FRMOPTIONSCOMPILER = pWindow->Create( hWndParent, "", @frmOptionsCompiler_WndProc, 0, 0, 0, 0, _
        WS_CHILD, WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   ' Height and width of this child form is set in frmOptions_OnNotify when the treeview option is selected.
   

   pWindow->AddControl("LABEL", , IDC_FRMOPTIONSCOMPILER_LBLTOOLCHAINS, _
        L(420,"Toolchains") & ":", 0, 2, 375, 15, _
        WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   
   dim as HWND hListBox = _
   pWindow->AddControl("LISTBOX", , IDC_FRMOPTIONSCOMPILER_LSTTOOLCHAINS, "", 0, 20, 386, 124, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or LBS_NOTIFY Or LBS_NOINTEGRALHEIGHT, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMOPTIONSCOMPILER_LBLSWITCHES, _
        L(153,"Additional compiler option switches (optional)" & ":"), 0, 148, 375, 16, _
        WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMOPTIONSCOMPILER_TXTFBSWITCHES, "", 0, 167, 386, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSCOMPILER_CHKRUNVIACOMMANDWINDOW, _
        L(273,"Run compiled programs using command window"), 0, 195, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSCOMPILER_CHKDISABLECOMPILEBEEP, _
        L(413,"Disable successful compile sound"), 0, 216, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("LABEL", , IDC_FRMOPTIONSCOMPILER_LBLFBHELP, _
        L(155,"FreeBASIC Help file (*.chm)") & ":", 0, 300, 375, 18, _
        WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMOPTIONSCOMPILER_TXTFBHELPFILE, "", 0, 320, 356, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("BUTTON", , IDC_FRMOPTIONSCOMPILER_CMDFBHELPFILE, "...", 362, 320, 24, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   
   gConfig.FBWINCompiler32  = ProcessFromCurdriveApp(gConfig.FBWINCompiler32)
   gConfig.FBWINCompiler64  = ProcessFromCurdriveApp(gConfig.FBWINCompiler64)
   gConfig.CompilerHelpfile = ProcessFromCurdriveApp(gConfig.CompilerHelpfile)
   gConfig.WinFBXHelpfile   = ProcessFromCurdriveApp(gConfig.WinFBXHelpfile)

   ' Load the Toolchain listbox
   DIM hSearch as HANDLE
   dim WFD AS WIN32_FIND_DATAW
   DIM wszPath AS WSTRING * MAX_PATH
   dim wszCurPath AS WSTRING * MAX_PATH
   dim wszFullPath AS WSTRING * (MAX_PATH * 2)
   dim wszDisplayName AS WSTRING * (MAX_PATH * 2)
   dim as long nCount = 0
   dim as long idx = 0
   dim as long idxMatch = -1
   
   wszPath = AfxGetExePathName + "toolchains\"
   wszCurPath = wszPath + "*"

   hSearch = FindFirstFile( wszCurPath, @WFD )
   IF hSearch <> INVALID_HANDLE_VALUE THEN
      DO
         IF (WFD.dwFileAttributes AND FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY THEN
            if (WFD.cFileName <> ".") andalso (WFD.cFileName <> "..") then
               wszFullPath = ucase( wszPath & WFD.cFileName & "\fbc32.exe" )

               ' Do check to see if path contains spaces as we will have to warn the user to fix it
               wszDisplayName = WFD.cFileName
               if instr( wszFullPath, " " ) then
                  wszDisplayName = wszDisplayName & "   (Warning: File path contains spaces)"
               end if

               idx = SendMessage( hListBox, LB_ADDSTRING, 0, cast(LPARAM, @wszDisplayName) )
               if ucase(gConfig.FBWINCompiler32) = wszFullPath then
                  idxMatch = idx
               end if
               
            end if   
         END IF
      LOOP WHILE FindNextFile( hSearch, @WFD )
      FindClose( hSearch )
   END IF
   SendMessage( hListBox, LB_SETCURSEL, idxMatch, 0 )
   
   Dim As HWnd HWnd = HWND_FRMOPTIONSCOMPILER
   AfxSetWindowText( GetDlgItem( HWnd, IDC_FRMOPTIONSCOMPILER_TXTFBSWITCHES),     gConfig.CompilerSwitches)
   AfxSetWindowText( GetDlgItem( HWnd, IDC_FRMOPTIONSCOMPILER_TXTFBHELPFILE),     gConfig.CompilerHelpfile)
   AfxSetWindowText( GetDlgItem( HWnd, IDC_FRMOPTIONSCOMPILER_TXTWINFBXHELPPATH), gConfig.WinFBXHelpfile)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSCOMPILER_CHKRUNVIACOMMANDWINDOW), gConfig.RunViaCommandWindow)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSCOMPILER_CHKDISABLECOMPILEBEEP),  gConfig.DisableCompileBeep)

   Function = 0 
End Function
