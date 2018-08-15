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
' frmOptionsEditor Window procedure
' ========================================================================================
Function frmOptionsEditor_WndProc( ByVal HWnd   As HWnd, _
                                   ByVal uMsg   As UINT, _
                                   ByVal wParam As WPARAM, _
                                   ByVal lParam As LPARAM _
                                   ) As LRESULT

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmOptionsEditor_Show
' ========================================================================================
Function frmOptionsEditor_Show( ByVal hWndParent As HWnd, _
                                ByVal nCmdShow   As Long = 0 _
                                ) As Long

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   HWND_FRMOPTIONSEDITOR = pWindow->Create( hWndParent, "", @frmOptionsEditor_WndProc, 0, 0, 0, 0, _
        WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   ' Height and width of this child form is set in frmOptions_OnNotify when the treeview option is selected.
   
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKSYNTAXHIGHLIGHTING, L(114,"Syntax highlighting"), 0, 0, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKCODETIPS, L(115,"Enable CodeTips"), 0, 21, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKAUTOCOMPLETE, L(95,"Enable Autocomplete"), 0, 42, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKCONFINECARET, L(116,"Confine caret to text"), 0, 63, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKCURRENTLINE, L(117,"Highlight current line"), 0, 84, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKTABTOSPACES, L(118,"Treat Tab as spaces"), 0, 105, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKAUTOINDENTATION, L(120,"Auto indentation"), 0, 126, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKSHOWLEFTMARGIN, L(122,"Show left margin"), 0, 147, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKSHOWFOLDMARGIN, L(123,"Show fold margin"), 0, 168, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKLINENUMBERING, L(124,"Show line numbers"), 0, 189, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKINDENTGUIDES, L(125,"Show Indentation guides"), 0, 210, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKSHOWRIGHTEDGE, L(306,"Show right edge column"), 0, 231, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMOPTIONSEDITOR_TXTRIGHTEDGE, "", 50, 252, 30, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL Or ES_NUMBER, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("LABEL", , IDC_FRMOPTIONSEDITOR_LBLRIGHTEDGE, L(307,"Position"), 85, 255, 100, 17, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
        
   pWindow->AddControl("TEXTBOX", , IDC_FRMOPTIONSEDITOR_TXTTABSIZE, "", 0, 320, 30, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL Or ES_NUMBER, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("COMBOBOX", , IDC_FRMOPTIONSEDITOR_COMBOCASE, "", 230, 319, 157, 22, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("LABEL", , IDC_FRMOPTIONSEDITOR_LBLTABSIZE, L(127,"Tab Size:"), 36, 323, 100, 17, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMOPTIONSEDITOR_LBLKEYWORDCASE, L(128,"Case:"), 120, 323, 100, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)


   Dim As Long idx = 0
   Dim As HWnd HWnd = HWND_FRMOPTIONSEDITOR
   
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKSYNTAXHIGHLIGHTING), gConfig.SyntaxHighlighting)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKCODETIPS),           gConfig.CodeTips)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKAUTOCOMPLETE),       gConfig.AutoComplete)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKSHOWLEFTMARGIN),     gConfig.LeftMargin)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKSHOWFOLDMARGIN),     gConfig.FoldMargin)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKCONFINECARET),       gConfig.ConfineCaret)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKLINENUMBERING),      gConfig.LineNumbering)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKINDENTGUIDES),       gConfig.IndentGuides)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKCURRENTLINE),        gConfig.HighlightCurrentLine)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKTABTOSPACES),        gConfig.TabIndentSpaces)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKAUTOINDENTATION),    gConfig.AutoIndentation)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKUNICODE),            gConfig.UnicodeEncoding)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKSHOWRIGHTEDGE),      gConfig.RightEdge)

   SetWindowTextW GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_TXTRIGHTEDGE), gConfig.RightEdgePosition
   SetWindowTextW GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_TXTTABSIZE), gConfig.TabSize
   
   ComboBox_AddString( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_COMBOCASE), @L(130,"Lower Case") )
   ComboBox_AddString( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_COMBOCASE), @L(131,"Upper Case") )
   ComboBox_AddString( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_COMBOCASE), @L(272,"Mixed Case") )
   ComboBox_AddString( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_COMBOCASE), @L(132,"Original Case") )
   ComboBox_SetCurSel( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_COMBOCASE), gConfig.KeywordCase )
   
   Function = 0
   
End Function



