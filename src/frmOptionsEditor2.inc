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

#include "frmOptionsEditor2.bi"
#include "clsConfig.bi"


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmOptions
' ========================================================================================
private Function frmOptionsEditor2_OnCommand( ByVal HWnd As HWnd, _
                                              ByVal id As Long, _
                                              ByVal hwndCtl As HWnd, _
                                              ByVal codeNotify As UINT _
                                              ) As LRESULT

   Select Case id
      
      case IDC_FRMOPTIONSEDITOR2_CHKAUTOINDENTATION
         if codeNotify = BN_CLICKED THEN
            if Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR2_CHKAUTOINDENTATION) ) then
               EnableWindow( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR2_CHKFORNEXTVARIABLE), true )
            else
               EnableWindow( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR2_CHKFORNEXTVARIABLE), false )
            end if   
         end if
   
   end select      

   function = 0
end function


' ========================================================================================
' frmOptionsEditor2 Window procedure
' ========================================================================================
private Function frmOptionsEditor2_WndProc( ByVal HWnd   As HWnd, _
                                            ByVal uMsg   As UINT, _
                                            ByVal wParam As WPARAM, _
                                            ByVal lParam As LPARAM _
                                            ) As LRESULT
   Select Case uMsg
      HANDLE_MSG (HWnd, WM_COMMAND,  frmOptionsEditor2_OnCommand)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmOptionsEditor2_Show
' ========================================================================================
public Function frmOptionsEditor2_Show( ByVal hWndParent As HWnd ) as LRESULT

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   HWND_FRMOPTIONSEDITOR2 = pWindow->Create( hWndParent, "", @frmOptionsEditor2_WndProc, 0, 0, 0, 0, _
        WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   ' Height and width of this child form is set in frmOptions_OnNotify when the treeview option is selected.
   
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR2_CHKCODETIPS, L(115,"Enable CodeTips"), 0, 0, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR2_CHKAUTOCOMPLETE, L(95,"Enable Autocomplete"), 0, 21, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR2_CHKCHARAUTOCOMPLETE, L(417,"Enable Character Autocompletion"), 0, 42, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR2_CHKAUTOINDENTATION, L(120,"Enable Auto Indentation"), 0, 63, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR2_CHKFORNEXTVARIABLE, L(416,"Append loop variable to For/Next statement"), 30, 84, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR2_CHKBRACEHIGHLIGHT, L(412,"Enable Brace Highlighting"), 0, 105, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR2_CHKOCCURRENCEHIGHLIGHT, L(419,"Enable Occurrences Highlighting"), 0, 126, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   Dim As HWnd HWnd = HWND_FRMOPTIONSEDITOR2
   
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR2_CHKCODETIPS),            gConfig.CodeTips)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR2_CHKAUTOCOMPLETE),        gConfig.AutoComplete)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR2_CHKCHARAUTOCOMPLETE),    gConfig.CharacterAutoComplete)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR2_CHKAUTOINDENTATION),     gConfig.AutoIndentation)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR2_CHKFORNEXTVARIABLE),     gConfig.ForNextVariable)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR2_CHKBRACEHIGHLIGHT),      gConfig.BraceHighlight)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR2_CHKOCCURRENCEHIGHLIGHT), gConfig.OccurrenceHighlight)

   if Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR2_CHKAUTOINDENTATION) ) then
      EnableWindow( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR2_CHKFORNEXTVARIABLE), true )
   else
      EnableWindow( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR2_CHKFORNEXTVARIABLE), false )
   end if   

   Function = 0
   
End Function



