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
' Save all options to the gConfig class
' ========================================================================================
Sub SaveEditorOptions()    
   
   Dim HWnd As HWnd 
   Dim wText As WString * MAX_PATH  
    
   ' GENERAL OPTIONS   
   HWnd = HWND_FRMOPTIONSGENERAL
   gConfig.MultipleInstances    = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSGENERAL_CHKMULTIPLEINSTANCES) )
   gConfig.CompileAutosave      = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSGENERAL_CHKCOMPILEAUTOSAVE) )
   gConfig.CloseFuncList        = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSGENERAL_CHKCLOSEFUNCLIST) )
   gConfig.HideCompile          = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSGENERAL_CHKHIDECOMPILE) )
   gConfig.AskExit              = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSGENERAL_CHKASKEXIT) )
   gConfig.HideToolbar          = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSGENERAL_CHKHIDETOOLBAR) )
   gConfig.HideStatusbar        = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSGENERAL_CHKHIDESTATUSBAR) )

   ' EDITOR OPTIONS   
   HWnd = HWND_FRMOPTIONSEDITOR
   gConfig.SyntaxHighlighting   = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_CHKSYNTAXHIGHLIGHTING) )
   gConfig.Codetips             = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_CHKCODETIPS) )
   gConfig.AutoComplete         = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_CHKAUTOCOMPLETE) )
   gConfig.LeftMargin           = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_CHKSHOWLEFTMARGIN) )
   gConfig.FoldMargin           = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_CHKSHOWFOLDMARGIN) )
   gConfig.ConfineCaret         = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_CHKCONFINECARET) )
   gConfig.LineNumbering        = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_CHKLINENUMBERING) )
   gConfig.HighlightCurrentLine = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_CHKCURRENTLINE) )
   gConfig.TabIndentSpaces      = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_CHKTABTOSPACES) )
   gConfig.IndentGuides         = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_CHKINDENTGUIDES ) )
   gConfig.AutoIndentation      = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_CHKAUTOINDENTATION) )
   gConfig.UnicodeEncoding      = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_CHKUNICODE) )
   gConfig.RightEdge            = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_CHKSHOWRIGHTEDGE) )

   gConfig.RightEdgePosition = AfxGetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_TXTRIGHTEDGE) )
   gConfig.TabSize = AfxGetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_TXTTABSIZE) )
   gConfig.KeywordCase = ComboBox_GetCurSel( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_COMBOCASE))
   
   ' If the localization has changed then display a message indicating that the changes
   ' will not be seen until the application is restarted.
   HWnd = HWND_FRMOPTIONSLOCAL
   wText = AfxGetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSLOCAL_LBLLOCALIZATION) )
   If Ucase(wText) <> Ucase(gConfig.LocalizationFile) Then
      MessageBox( HWND_FRMMAIN, L(108,"Localization changes will be applied the next time the application is run."), _
                  L(109,"Warning"), MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 Or MB_APPLMODAL )
   End If
   gConfig.LocalizationFile = wText
   
   
   ' COLOR/FONT OPTIONS
   HWnd = HWND_FRMOPTIONSCOLORS
   ' Copy the temporary items to the main array  
   redim gConfig.Themes(ubound(gConfig.ThemesTemp))
   for i as long = lbound(gConfig.ThemesTemp) to ubound(gConfig.ThemesTemp)
      gConfig.Themes(i) = gConfig.ThemesTemp(i)   
   NEXT
   dim as hwnd hWndCombo = GetDlgItem(hwnd, IDC_FRMOPTIONSCOLORS_COMBOTHEMES)
   Dim As Long idxTheme = Combobox_GetCurSel( hWndCombo )
   gConfig.SelectedTheme = gConfig.ThemesTemp(idxTheme).id
   erase gConfig.ThemesTemp

   Dim idx As Long 
   idx = ComboBox_GetCurSel( GetDlgItem(HWnd, IDC_FRMOPTIONSCOLORS_COMBOFONTNAME) )
   gConfig.EditorFontname = AfxGetComboBoxText( GetDlgItem(HWnd, IDC_FRMOPTIONSCOLORS_COMBOFONTNAME), idx )
   idx = ComboBox_GetCurSel( GetDlgItem(HWnd, IDC_FRMOPTIONSCOLORS_COMBOFONTSIZE) )
   gConfig.EditorFontSize = AfxGetComboBoxText( GetDlgItem(HWnd, IDC_FRMOPTIONSCOLORS_COMBOFONTSIZE), idx )  
   idx = ComboBox_GetCurSel( GetDlgItem(HWnd, IDC_FRMOPTIONSCOLORS_COMBOFONTCHARSET) )
   gConfig.EditorFontCharSet = AfxGetComboBoxText( GetDlgItem(HWnd, IDC_FRMOPTIONSCOLORS_COMBOFONTCHARSET), idx )
   
   ' COMPILER OPTIONS
   HWnd = HWND_FRMOPTIONSCOMPILER
   gConfig.FBWINCompiler32  = AfxGetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSCOMPILER_TXTFBWIN32) )
   gConfig.FBWINCompiler64  = AfxGetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSCOMPILER_TXTFBWIN64) )
   gConfig.CompilerSwitches = AfxGetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSCOMPILER_TXTFBSWITCHES) )
   gConfig.CompilerHelpfile = AfxGetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSCOMPILER_TXTFBHELPFILE) )
   gConfig.Win32APIHelpfile = AfxGetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSCOMPILER_TXTWIN32HELPPATH) )
   gConfig.WinFBXHelpfile   = AfxGetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSCOMPILER_TXTWINFBXHELPPATH) )
   gConfig.RunViaCommandWindow = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSCOMPILER_CHKRUNVIACOMMANDWINDOW) )
   
   
   ' KEYWORDS
   If gConfig.bKeywordsDirty Then
      HWnd = HWND_FRMOPTIONSKEYWORDS
      Dim As HWnd hTextBox = GetDlgItem(HWnd, IDC_FRMOPTIONSKEYWORDS_TXTKEYWORDS)
      Dim As Long nLen = GetWindowTextLength(hTextBox) + 2  ' add space for ending nuls
      Dim pzswKeywords As WString Ptr = CAllocate(nLen*2)
      GetWindowText( hTextBox, pzswKeywords, nLen )
      Dim As String FBKeywords = Str(*pzswKeywords)
      gConfig.FBKeywords = Trim(FBKeywords)
      gConfig.SaveKeywords
      Deallocate(pzswKeywords)
   End If
   
   ' Output the config settings to disk file
   gConfig.SaveConfigFile

End Sub


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmOptions
' ========================================================================================
Function frmOptions_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN

   ' This is a modal popup window so disable the parent window
   EnableWindow GetParent(HWnd), False

   '  Message cracker macro expects a True to be returned for a successful
   '  OnCreate handler even though returning -1 from a standard WM_CREATE
   '  call would stop creating the window. This is just one of those Windows
   '  inconsistencies.
   Return True
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmOptions
' ========================================================================================
Function frmOptions_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Select Case id
      Case IDOK
         If codeNotify = BN_CLICKED Then
            SaveEditorOptions()
            ' Reposition main windows in case hide toolbar/statusbar selected
            frmMain_PositionWindows(HWND_FRMMAIN)
            ' Apply the newly saved options to all open Scintilla windows
            Dim pDoc As clsDocument Ptr
            for i as long = lbound(gApp.Projects) to ubound(gApp.Projects)
               if gApp.Projects(i).InUse THEN
                  For y as long = 0 To gApp.Projects(i).GetDocumentCount - 1
                     pDoc = gApp.Projects(i).GetDocumentPtr(y)
                     If pDoc Then pDoc->ApplyProperties
                  Next
               end if
            next   
            SendMessage HWnd, WM_CLOSE, 0, 0
            Exit Function
         End If
      Case IDCANCEL
         If codeNotify = BN_CLICKED Then
            SendMessage HWnd, WM_CLOSE, 0, 0
            Exit Function
         End If
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmOptions
' ========================================================================================
Function frmOptions_OnNotify(ByVal HWnd As HWnd, ByVal id As Long, ByVal pNMHDR As NMHDR Ptr) As LRESULT

    Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
    
    Select Case id
       Case IDC_FRMOPTIONS_TVWCATEGORIES
          If pNMHDR->code = TVN_SELCHANGED Then
             Dim As Long nLeft, nTop
             Dim As HWnd hForm, hLabel

             Static hCurrent As HWnd   ' hwnd of currently display child Form
             
             If hCurrent Then ShowWindow hCurrent, SW_HIDE
             
             Dim lpNMTV As NM_TREEVIEW Ptr = Cast(NM_TREEVIEW Ptr, pNMHDR)
             hLabel = GetDlgItem(HWnd, IDC_FRMOPTIONS_LBLCATEGORY)
             
             ' Display the correct Category name and Child dialog. The child identifier is held in the lParam
             Select Case FF_TreeView_GetlParam( pNMHDR->hWndFrom, lpNMTV->itemNew.hItem)
                 Case 0
                     SetWindowTextW( hLabel, WStr(" ") & L(226,"General Options")) 
                     hForm = HWND_FRMOPTIONSGENERAL
                 Case 1
                     SetWindowTextW( hLabel, WStr(" ") & L(110,"Code Editor")) 
                     hForm = HWND_FRMOPTIONSEDITOR
                 Case 2
                     SetWindowTextW( hLabel, WStr(" ") & L(111,"Colors and Fonts") )
                     hForm = HWND_FRMOPTIONSCOLORS
                 Case 3
                     SetWindowTextW( hLabel, WStr(" ") & L(112,"Compiler Setup") )
                     hForm = HWND_FRMOPTIONSCOMPILER
                 Case 4
                     SetWindowTextW( hLabel, WStr(" ") & L(129,"Localization") )
                     hForm = HWND_FRMOPTIONSLOCAL
                 Case 5
                     SetWindowTextW( hLabel, WStr(" ") & L(177,"FreeBASIC Keywords") )
                     hForm = HWND_FRMOPTIONSKEYWORDS
             End Select
             
         
             ' Move the child Form into place. The child form will be displayed under the header label.
             Dim rc As Rect
             GetWindowRect( hLabel, @rc )
             MapWindowPoints( Null, HWnd, Cast(LPPOINT, @rc), 2 )
             SetWindowPos hForm, HWND_TOP, _
                            rc.Left + pWindow->ScaleX(5), pWindow->ScaleY(50), _
                            pWindow->ScaleX(420), pWindow->ScaleY(350), _
                            SWP_SHOWWINDOW
             
             InvalidateRect HWnd, ByVal Null, True
             UpdateWindow HWnd
             
             hCurrent = hForm
          End If
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CTLCOLORSTATIC message for window/dialog: frmOptions
' ========================================================================================
Function frmOptions_OnCtlColorStatic(HWnd As HWnd, hdc As HDC, hWndChild As HWnd, nType As Long) As HBRUSH

   Select Case hWndChild
      Case GetDlgItem( HWnd, IDC_FRMOPTIONS_LBLCATEGORY )
         ' Set the category label to blue text 
         SetTextColor hdc, 8388608
         SetBkColor   hdc, GetSysColor( COLOR_BTNFACE )
         Function = GetSysColorBrush( COLOR_BTNFACE )
   End Select
         
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmOptions
' ========================================================================================
Function frmOptions_OnClose(HWnd As HWnd) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableWindow GetParent(HWnd), True 
   DestroyWindow HWnd
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmOptions
' ========================================================================================
Function frmOptions_OnDestroy(HWnd As HWnd) As LRESULT
   Dim As HFONT hFont = AfxGetWindowFont(GetDlgItem(HWnd, IDC_FRMOPTIONS_LBLCATEGORY))
   DeleteFont(hFont)

   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' frmOptions Window procedure
' ========================================================================================
Function frmOptions_WndProc( ByVal HWnd   As HWnd, _
                             ByVal uMsg   As UINT, _
                             ByVal wParam As WPARAM, _
                             ByVal lParam As LPARAM _
                             ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,   frmOptions_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,    frmOptions_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmOptions_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,  frmOptions_OnCommand)
      HANDLE_MSG (HWnd, WM_NOTIFY,   frmOptions_OnNotify)
      HANDLE_MSG (HWnd, WM_CTLCOLORSTATIC, frmOptions_OnCtlColorStatic)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmOptions_Show
' ========================================================================================
Function frmOptions_Show( ByVal hWndParent As HWnd, _
                          ByVal nCmdShow   As Long = 0 _
                          ) As Long

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   pWindow->Create( hWndParent, L(113,"Environment Options"), @frmOptions_WndProc, 0, 0, 0, 0, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT )
   pWindow->SetClientSize(625, 460)
   pWindow->Center(pWindow->hWindow, hWndParent)
      
   Dim As HWnd hLabel = _
      pWindow->AddControl("LABEL", , IDC_FRMOPTIONS_LBLCATEGORY, "Category", 216, 10, 400, 30, _
           WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_CENTERIMAGE Or SS_NOTIFY Or SS_SUNKEN, _
           WS_EX_LEFT Or WS_EX_LTRREADING)
   Dim As HFONT hFont = pWindow->CreateFont("", 12, FW_BOLD)
   AfxSetWindowFont hLabel, hFont, True

   Dim As HWnd hWndTreeview = _
        pWindow->AddControl("TREEVIEW", , IDC_FRMOPTIONS_TVWCATEGORIES, "", 10, 10, 200, 385, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or TVS_HASBUTTONS Or TVS_HASLINES Or TVS_LINESATROOT Or TVS_SHOWSELALWAYS, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR)
   
   pWindow->AddControl("LABEL", , IDC_FRMOPTIONS_LABEL1, "", 8, 405, 600, 3, _
        WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_NOTIFY Or SS_SUNKEN, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("BUTTON", , IDOK, L(0,"&OK"), 450, 417, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_DEFPUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   
   pWindow->AddControl("BUTTON", , IDCANCEL, L(1,"&Cancel"), 535, 417, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)


   ' Configure the TreeView
   Dim hItem As HANDLE 
   hItem = FF_TreeView_InsertItem (hWndTreeview, TVI_ROOT, L(226,"General Options"),    0)
   FF_TreeView_InsertItem (hWndTreeview, TVI_ROOT, L(110,"Code Editor"),        1)
   FF_TreeView_InsertItem (hWndTreeview, TVI_ROOT, L(111,"Colors and Fonts"),   2)
   FF_TreeView_InsertItem (hWndTreeview, TVI_ROOT, L(112,"Compiler Setup"),     3)
   FF_TreeView_InsertItem (hWndTreeview, TVI_ROOT, L(129,"Localization"),       4)
   FF_TreeView_InsertItem (hWndTreeview, TVI_ROOT, L(177,"FreeBASIC Keywords"), 5)

   ' Load all of the child dialogs
   frmOptionsGeneral_Show( pWindow->hWindow )
   frmOptionsEditor_Show( pWindow->hWindow )
   frmOptionsColors_Show( pWindow->hWindow )
   frmOptionsCompiler_Show( pWindow->hWindow )
   frmOptionsLocal_Show( pWindow->hWindow )
   frmOptionsKeywords_Show( pWindow->hWindow )

   ' Select the first item "Code Editor"
   TreeView_SelectItem( hWndTreeview, hItem)

   SetFocus hWndTreeview
   
   ' Process Windows messages
   Function = pWindow->DoEvents(nCmdShow)
   
   ' Delete the frmOptions CWindow class manually allocated memory 
   Delete pWindow

   ' Delete the child windows releasing their manually allocated memory 
   pWindow = AfxCWindowPtr(HWND_FRMOPTIONSGENERAL):   Delete pWindow
   pWindow = AfxCWindowPtr(HWND_FRMOPTIONSEDITOR):    Delete pWindow
   pWindow = AfxCWindowPtr(HWND_FRMOPTIONSCOLORS):    Delete pWindow
   pWindow = AfxCWindowPtr(HWND_FRMOPTIONSCOMPILER):  Delete pWindow
   pWindow = AfxCWindowPtr(HWND_FRMOPTIONSLOCAL):     Delete pWindow
   pWindow = AfxCWindowPtr(HWND_FRMOPTIONSKEYWORDS):  Delete pWindow

End Function





