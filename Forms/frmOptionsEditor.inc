

' ========================================================================================
' frmOptionsEditor Window procedure
' ========================================================================================
Function frmOptionsEditor_WndProc( ByVal HWnd   As HWnd, _
                                   ByVal uMsg   As UINT, _
                                   ByVal wParam As WPARAM, _
                                   ByVal lParam As LPARAM _
                                   ) As LRESULT

   ' for messages that we don't deal with
   Function = DefWindowProcW(HWnd, uMsg, wParam, lParam)

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
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKCONFINECARET, L(116,"Confine caret to text"), 0, 42, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKCURRENTLINE, L(117,"Highlight current line"), 0, 63, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKTABTOSPACES, L(118,"Treat Tab as spaces"), 0, 84, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKMULTIPLEINSTANCES, L(119,"Allow multiple instances"), 0, 105, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKAUTOINDENTATION, L(120,"Auto indentation"), 0, 126, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKCOMPILEAUTOSAVE, L(121,"Autosave files before compiling"), 0, 147, 235, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKSHOWLEFTMARGIN, L(122,"Show left margin"), 0, 168, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKSHOWFOLDMARGIN, L(123,"Show fold margin"), 0, 189, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKLINENUMBERING, L(124,"Show line numbers"), 0, 210, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKINDENTGUIDES, L(125,"Show Indentation guides"), 0, 231, 400, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMOPTIONSEDITOR_TXTTABSIZE, "", 2, 270, 30, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL Or ES_NUMBER, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("COMBOBOX", , IDC_FRMOPTIONSEDITOR_COMBOCASE, "", 230, 269, 157, 22, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("LABEL", , IDC_FRMOPTIONSEDITOR_LBLTABSIZE, L(127,"Tab Size:"), 38, 273, 100, 17, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMOPTIONSEDITOR_LBLKEYWORDCASE, L(128,"Case:"), 120, 273, 100, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)


   Dim As Long idx = 0
   Dim As HWnd HWnd = HWND_FRMOPTIONSEDITOR
   
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKSYNTAXHIGHLIGHTING), gConfig.SyntaxHighlighting)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKCODETIPS),           gConfig.CodeTips)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKSHOWLEFTMARGIN),     gConfig.LeftMargin)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKSHOWFOLDMARGIN),     gConfig.FoldMargin)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKCONFINECARET),       gConfig.ConfineCaret)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKLINENUMBERING),      gConfig.LineNumbering)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKINDENTGUIDES),       gConfig.IndentGuides)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKCURRENTLINE),        gConfig.HighlightCurrentLine)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKTABTOSPACES),        gConfig.TabIndentSpaces)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKMULTIPLEINSTANCES),  gConfig.MultipleInstances)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKCOMPILEAUTOSAVE),    gConfig.CompileAutoSave)
   Button_SetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_CHKAUTOINDENTATION),    gConfig.AutoIndentation)

   SetWindowTextW GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_TXTTABSIZE), @gConfig.TabSize
   
   ComboBox_AddString( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_COMBOCASE), @L(130,"Lower Case") )
   ComboBox_AddString( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_COMBOCASE), @L(131,"Upper Case") )
   ComboBox_AddString( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_COMBOCASE), @L(132,"Original Case") )
   ComboBox_SetCurSel( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_COMBOCASE), gConfig.KeywordCase )
   

   Function = 0
   
End Function
