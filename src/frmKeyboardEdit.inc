'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2022 Paul Squires, PlanetSquires Software
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

#include once "frmKeyboard.bi"


' ========================================================================================
' Generate the keybind string based on the selected controls
' ========================================================================================
function frmKeyboardEdit_GenerateKeyBinding() as CWSTR
   dim as HWND hCtrl
   dim as CWSTR wszKeys
   dim as CWSTR wszAccel

   hCtrl = GetDlgItem( HWND_FRMKEYBOARDEDIT, IDC_FRMKEYBOARDEDIT_CHKCTRL )
   if Button_GetCheck(hCtrl) = BST_CHECKED then wszKeys = "Ctrl+"
   hCtrl = GetDlgItem( HWND_FRMKEYBOARDEDIT, IDC_FRMKEYBOARDEDIT_CHKALT )
   if Button_GetCheck(hCtrl) = BST_CHECKED then wszKeys = wszKeys & "Alt+"
   hCtrl = GetDlgItem( HWND_FRMKEYBOARDEDIT, IDC_FRMKEYBOARDEDIT_CHKSHIFT )
   if Button_GetCheck(hCtrl) = BST_CHECKED then wszKeys = wszKeys & "Shift+"

   hCtrl = GetDlgItem( HWND_FRMKEYBOARDEDIT, IDC_FRMKEYBOARDEDIT_COMBOACCEL )
   dim as long nCursel = ComboBox_GetCurSel( hCtrl )
   wszAccel = AfxGetComboBoxText( hCtrl, nCurSel )
   if wszAccel = "None" then 
      wszKeys = ""
   else   
      wszKeys = wszKeys & wszAccel
   end if   

   return wszKeys
end function

' ========================================================================================
' Process WM_CREATE message for window/dialog: frmKeyboardEdit
' ========================================================================================
Function frmKeyboardEdit_OnCreate( _
            ByVal HWnd As HWnd, _
            ByVal lpCreateStructPtr As LPCREATESTRUCT _
            ) As BOOLEAN

   ' This is a modal popup window so disable the parent window
   DisableAllModeless()
   EnableWindow( GetParent(Hwnd), false )

   '  Message cracker macro expects a True to be returned for a successful
   '  OnCreate handler even though returning -1 from a standard WM_CREATE
   '  call would stop creating the window. This is just one of those Windows
   '  inconsistencies.
   Return True
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmKeyboard
' ========================================================================================
Function frmKeyboardEdit_OnCommand( _
            ByVal HWnd As HWnd, _
            ByVal id As Long, _
            ByVal hwndCtl As HWnd, _
            ByVal codeNotify As UINT _
            ) As LRESULT

   Select Case id

      case IDC_FRMKEYBOARDEDIT_CHKCTRL, IDC_FRMKEYBOARDEDIT_CHKALT, IDC_FRMKEYBOARDEDIT_CHKSHIFT
         if codeNotify = BN_CLICKED then
            Dim As Long nCurSel = ListView_GetSelection( HWND_FRMKEYBOARD_LISTVIEW )
            If nCurSel < 0 Then exit function
            dim as CWSTR wszKeys = frmKeyboardEdit_GenerateKeyBinding()
            frmKeyboard_CheckForKeyConflict( wszKeys, nCurSel )
         end if
          
      case IDC_FRMKEYBOARDEDIT_COMBOACCEL
         if codeNotify = CBN_SELCHANGE then
            Dim As Long nCurSel = ListView_GetSelection( HWND_FRMKEYBOARD_LISTVIEW )
            If nCurSel < 0 Then exit function
            dim as CWSTR wszKeys = frmKeyboardEdit_GenerateKeyBinding()
            frmKeyboard_CheckForKeyConflict( wszKeys, nCurSel )
         end if

      Case IDOK
         If codeNotify = BN_CLICKED Then
            ' Update the gKeysEdit entry with the new UserKeys value & Enabled value
            dim as CWSTR wszKeys = frmKeyboardEdit_GenerateKeyBinding()
            gKeysEdit.wszUserKeys = wszKeys
            dim as HWND hCtrl = GetDlgItem( HWND_FRMKEYBOARDEDIT, IDC_FRMKEYBOARDEDIT_CHKDISABLED )
            gKeysEdit.bDefaultDisabled = iif( Button_GetCheck(hCtrl) = BST_CHECKED, true, false )
            SendMessage( HWnd, WM_CLOSE, 0, 0 )
            Exit Function
         end if
            
      Case IDCANCEL
         If codeNotify = BN_CLICKED Then
            SendMessage( HWnd, WM_CLOSE, 0, 0 )
            Exit Function
         End If
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmKeyboardEdit
' ========================================================================================
Function frmKeyboardEdit_OnClose( byval HWnd As HWnd ) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableAllModeless()
   EnableWindow( GetParent(Hwnd), true )
   DestroyWindow( HWnd )
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmKeyboardEdit
' ========================================================================================
Function frmKeyboardEdit_OnDestroy( byval HWnd As HWnd) As LRESULT
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' frmKeyboard Window procedure
' ========================================================================================
Function frmKeyboardEdit_WndProc( _
            ByVal HWnd   As HWnd, _
            ByVal uMsg   As UINT, _
            ByVal wParam As WPARAM, _
            ByVal lParam As LPARAM _
            ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,      frmKeyboardEdit_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,       frmKeyboardEdit_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,     frmKeyboardEdit_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,     frmKeyboardEdit_OnCommand)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc( HWnd, uMsg, wParam, lParam )

End Function


' ========================================================================================
' frmKeyboardEdit_Show
' ========================================================================================
Function frmKeyboardEdit_Show( ByVal hWndParent As HWnd ) As LRESULT

   dim hCtrl as HWnd
   dim as CWSTR wszText
   
   ' Get the components of the current user key binding
   dim as CWSTR wszKeys = gKeysEdit.wszUserKeys
   dim as long nCount = AfxStrParseCount( wszKeys, "+" )
   dim as CWSTR wszAccel
   dim as long isAlt = BST_UNCHECKED
   dim as long isShift = BST_UNCHECKED
   dim as long isCtrl = BST_UNCHECKED
   dim as long isDisabled = BST_UNCHECKED
   
   if gKeysEdit.bDefaultDisabled then
      isDisabled = BST_CHECKED
   end if   
   
   for ii as long = nCount to 1 step -1
      ' The accelerator key will always be the last parse regardless of 
      ' number of items in the string
      dim as CWSTR wszString = AfxStrParse( wszKeys, ii, "+" )
      if ii = nCount then
         wszAccel = wszString
      else
         select case ucase(wszString)
            case "CTRL": isCtrl = BST_CHECKED
            case "SHIFT": isShift = BST_CHECKED
            case "ALT": isAlt = BST_CHECKED
         end select
      end if
   next
   
   
   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   HWND_FRMKEYBOARDEDIT = _
   pWindow->Create(hWndParent, L(220,"Keyboard Shortcuts"), _
        @frmKeyboardEdit_WndProc, 0, 0, 0, 0, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT )
   pWindow->SetClientSize( 330, 220 )
   pWindow->Center( pWindow->hWindow, hWndParent )

   hCtrl = pWindow->AddControl("LABEL", , IDC_FRMKEYBOARDEDIT_LABEL1, "", 20, 10, 350, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   AfxSetWindowText( hCtrl, gKeysEdit.wszDescription )
   AfxSetWindowFont( hCtrl, ghStatusBar.hFontStatusBar, true )

   hCtrl = pWindow->AddControl("LABEL", , IDC_FRMKEYBOARDEDIT_LABEL2, "", 20, 30, 350, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszText = gKeysEdit.wszDefaultKeys
   if len(wszText) = 0 then wszText = "(no key binding)"
   AfxSetWindowText( hCtrl, "Default: " & wszText  )
   AfxSetWindowFont( hCtrl, ghStatusBar.hFontStatusBar, true )

   hCtrl = pWindow->AddControl("CHECKBOX", , IDC_FRMKEYBOARDEDIT_CHKDISABLED, "Disable Default", 20, 50, 150, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   AfxSetWindowFont( hCtrl, ghStatusBar.hFontStatusBar, true )
   Button_SetCheck( hCtrl, isDisabled )

   hCtrl = pWindow->AddControl("LABEL", , IDC_FRMKEYBOARDEDIT_LABEL3, "", 20, 94, 350, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   AfxSetWindowText( hCtrl, "User defined Keybinding" )
   AfxSetWindowFont( hCtrl, ghStatusBar.hFontStatusBar, true )

   hCtrl = pWindow->AddControl("CHECKBOX", , IDC_FRMKEYBOARDEDIT_CHKCTRL, "Ctrl", 20, 114, 48, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   AfxSetWindowFont( hCtrl, ghStatusBar.hFontStatusBar, true )
   Button_SetCheck( hCtrl, isCtrl )
   
   hCtrl = pWindow->AddControl("CHECKBOX", , IDC_FRMKEYBOARDEDIT_CHKALT, "Alt", 75, 114, 42, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   AfxSetWindowFont( hCtrl, ghStatusBar.hFontStatusBar, true )
   Button_SetCheck( hCtrl, isAlt )
   
   hCtrl = pWindow->AddControl("CHECKBOX", , IDC_FRMKEYBOARDEDIT_CHKSHIFT, "Shift", 127, 114, 48, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   AfxSetWindowFont( hCtrl, ghStatusBar.hFontStatusBar, true )
   Button_SetCheck( hCtrl, isShift )

   dim as HWND hCombo 
   hCombo = pWindow->AddControl("COMBOBOX", , IDC_FRMKEYBOARDEDIT_COMBOACCEL, "", 190, 114, 120, 200, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   AfxSetWindowFont( hCtrl, ghStatusBar.hFontStatusBar, true )

   pWindow->AddControl("BUTTON", , IDOK, L(0,"OK"), 154, 170, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDCANCEL, L(1,"Cancel"), 236, 170, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
                                     
   ComboBox_AddString( hCombo, @wstr("None") ) 
   ComboBox_AddString( hCombo, @wstr("Backspace") ) 
   ComboBox_AddString( hCombo, @wstr("Tab") ) 
   ComboBox_AddString( hCombo, @wstr("Enter") ) 
   ComboBox_AddString( hCombo, @wstr("Escape") ) 
   ComboBox_AddString( hCombo, @wstr("Spacebar") ) 
   ComboBox_AddString( hCombo, @wstr("PageUp") ) 
   ComboBox_AddString( hCombo, @wstr("PageDn") ) 
   ComboBox_AddString( hCombo, @wstr("End") ) 
   ComboBox_AddString( hCombo, @wstr("Home") ) 
   ComboBox_AddString( hCombo, @wstr("Left") ) 
   ComboBox_AddString( hCombo, @wstr("Right") ) 
   ComboBox_AddString( hCombo, @wstr("Up") ) 
   ComboBox_AddString( hCombo, @wstr("Down") ) 
   ComboBox_AddString( hCombo, @wstr("DEL") ) 
   ComboBox_AddString( hCombo, @wstr("INS") ) 
   for i as long = 48 to 57    ' 0 to 9
      wszText = wchr(i)
      ComboBox_AddString( hCombo, wszText.sptr ) 
   next
   for i as long = 65 to 90    ' A to Z
      wszText = wchr(i)
      ComboBox_AddString( hCombo, wszText.sptr ) 
   next
   ComboBox_AddString( hCombo, @wstr("Numpad0") ) 
   ComboBox_AddString( hCombo, @wstr("Numpad1") ) 
   ComboBox_AddString( hCombo, @wstr("Numpad2") ) 
   ComboBox_AddString( hCombo, @wstr("Numpad3") ) 
   ComboBox_AddString( hCombo, @wstr("Numpad4") ) 
   ComboBox_AddString( hCombo, @wstr("Numpad5") ) 
   ComboBox_AddString( hCombo, @wstr("Numpad6") ) 
   ComboBox_AddString( hCombo, @wstr("Numpad7") ) 
   ComboBox_AddString( hCombo, @wstr("Numpad8") ) 
   ComboBox_AddString( hCombo, @wstr("Numpad9") ) 
   ComboBox_AddString( hCombo, @wstr("Numpad*") ) 
   ComboBox_AddString( hCombo, @wstr("Numpad+") ) 
   ComboBox_AddString( hCombo, @wstr("Numpad-") ) 
   ComboBox_AddString( hCombo, @wstr("Numpad.") ) 
   ComboBox_AddString( hCombo, @wstr("Numpad/") ) 
   ComboBox_AddString( hCombo, @wstr("F1") ) 
   ComboBox_AddString( hCombo, @wstr("F2") ) 
   ComboBox_AddString( hCombo, @wstr("F3") ) 
   ComboBox_AddString( hCombo, @wstr("F4") ) 
   ComboBox_AddString( hCombo, @wstr("F5") ) 
   ComboBox_AddString( hCombo, @wstr("F6") ) 
   ComboBox_AddString( hCombo, @wstr("F7") ) 
   ComboBox_AddString( hCombo, @wstr("F8") ) 
   ComboBox_AddString( hCombo, @wstr("F9") ) 
   ComboBox_AddString( hCombo, @wstr("F10") ) 
   ComboBox_AddString( hCombo, @wstr("F11") ) 
   ComboBox_AddString( hCombo, @wstr("F12") ) 
   ComboBox_AddString( hCombo, @wstr("Tilde") ) 
   ComboBox_AddString( hCombo, @wstr("Plus") ) 
   ComboBox_AddString( hCombo, @wstr("Comma") ) 
   ComboBox_AddString( hCombo, @wstr("Minus") ) 
   ComboBox_AddString( hCombo, @wstr("Period") ) 
   ComboBox_AddString( hCombo, @wstr("SingleQuote") ) 
   ComboBox_AddString( hCombo, @wstr("SemiColon") ) 
   ComboBox_AddString( hCombo, @wstr("/") ) 
   ComboBox_AddString( hCombo, @wstr("\") ) 
   ComboBox_AddString( hCombo, @wstr("[") ) 
   ComboBox_AddString( hCombo, @wstr("]") ) 

   if len(wszAccel) = 0 then wszAccel = "None"
   dim as long nCurSel 
   nCursel = ComboBox_FindStringExact( hCombo, -1, wszAccel.sptr )
   ComboBox_SetCurSel( hCombo, nCursel )

   SetFocus( GetDlgItem(HWND_FRMKEYBOARDEDIT, IDOK) )    
   
   ' Process Windows messages(modal)
   Function = pWindow->DoEvents(SW_SHOW)
   
   ' Delete the CWindow class manually allocated memory 
   Delete pWindow

End Function

