
' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmOptionsEditor
' ========================================================================================
Function frmOptionsEditor_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Select Case id

      Case IDC_FRMOPTIONSEDITOR_CMDLOCALIZATION

         If codeNotify = BN_CLICKED Then
            ' Display the Open File Dialog
            Dim pwszName As WString Ptr = AfxIFileOpenDialogW(HWnd, id)
            If pwszName Then
               Dim wzFilename As WString * MAX_PATH = *pwszName
               FF_PATHNAME( WStr("NAMEX"), pwszName, @wzFilename )
               If LoadLocalizationFile(pwszName) Then
                  SetWindowTextW( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_LBLLOCALIZATION), wzFilename )
               End If
               CoTaskMemFree pwszName
            End If
            Exit Function
         End If

   End Select

   Function = 0
End Function


' ========================================================================================
' frmOptionsEditor Window procedure
' ========================================================================================
Function frmOptionsEditor_WndProc( ByVal HWnd   As HWnd, _
                                   ByVal uMsg   As UINT, _
                                   ByVal wParam As WPARAM, _
                                   ByVal lParam As LPARAM _
                                   ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_COMMAND, frmOptionsEditor_OnCommand)
   End Select

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

   HWND_FRMOPTIONSEDITOR = pWindow->Create( hWndParent, "", @frmOptionsEditor_WndProc, 0, 0, 354, 319, _
        WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->SetClientSize(354, 319)
   
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKSYNTAXHIGHLIGHTING, "Syntax highlighting", 0, 0, 179, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKCODETIPS, "Enable CodeTips", 0, 21, 179, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKCONFINECARET, "Confine caret to text", 0, 42, 179, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKCURRENTLINE, "Highlight current line", 0, 63, 179, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKTABTOSPACES, "Treat Tab as spaces", 0, 83, 179, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKMULTIPLEINSTANCES, "Allow multiple instances", 0, 104, 179, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKAUTOINDENTATION, "Auto indentation", 0, 124, 179, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKCOMPILEAUTOSAVE, "Autosave files before compiling", 0, 144, 235, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKSHOWLEFTMARGIN, "Show left margin", 185, 0, 179, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKSHOWFOLDMARGIN, "Show fold margin", 185, 21, 179, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKLINENUMBERING, "Show line numbers", 185, 42, 179, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSEDITOR_CHKINDENTGUIDES, "Show Indentation guides", 185, 64, 179, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMOPTIONSEDITOR_CMDLOCALIZATION, "Select", 255, 214, 69, 26, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMOPTIONSEDITOR_TXTTABSIZE, "", 2, 270, 30, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL Or ES_NUMBER, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("COMBOBOX", , IDC_FRMOPTIONSEDITOR_COMBOCASE, "", 208, 269, 125, 22, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("LABEL", , IDC_FRMOPTIONSEDITOR_LBLTABSIZE, "Tab Size:", 38, 273, 86, 17, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMOPTIONSEDITOR_LBLKEYWORDCASE, "Case:", 167, 273, 34, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMOPTIONSEDITOR_LBLLOCALIZATION, "English.lang", 14, 217, 206, 22, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_CENTER Or SS_CENTERIMAGE Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("GROUPBOX", , IDC_FRMOPTIONSEDITOR_FRAMELOCALIZATION, "Localization", 0, 198, 333, 50, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)


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
   SetWindowTextW GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_LBLLOCALIZATION), @gConfig.LocalizationFile
   SetWindowTextW GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_TXTTABSIZE), @gConfig.TabSize
   
   ComboBox_AddString( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_COMBOCASE), @wstr("Lower Case") )
   ComboBox_AddString( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_COMBOCASE), @wstr("Upper Case") )
   ComboBox_AddString( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_COMBOCASE), @wstr("Original Case") )
   
   idx = SendMessageW( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_COMBOCASE), CB_FINDSTRING, 0, _
                               Cast(LPARAM, Strptr(gConfig.KeywordCase)))
   ComboBox_SetCurSel( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_COMBOCASE), idx )
   

   Function = 0
   
End Function
