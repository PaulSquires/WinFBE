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

#include once "frmOptions.bi"
#include once "frmOptionsGeneral.bi"
#include once "frmOptionsEditor.bi"
#include once "frmOptionsEditor2.bi"
#include once "frmOptionsCompiler.bi"
#include once "frmOptionsColors.bi"
#include once "frmOptionsLocal.bi"
#include once "frmOptionsKeywords.bi"
#include once "frmOptionsKeywordsWinApi.bi"
#include once "frmMain.bi"

#include once "clsConfig.bi"


' ========================================================================================
' Save all options to the gConfig class
' ========================================================================================
private function frmOptions_SaveEditorOptions() as long   
   
   Dim HWnd As HWnd 
   Dim wText As WString * MAX_PATH  
    
   ' GENERAL OPTIONS   
   HWnd = HWND_FRMOPTIONSGENERAL
   gConfig.MultipleInstances    = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSGENERAL_CHKMULTIPLEINSTANCES) )
   gConfig.CompileAutosave      = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSGENERAL_CHKCOMPILEAUTOSAVE) )
   gConfig.AskExit              = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSGENERAL_CHKASKEXIT) )
   gConfig.RestoreSession       = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSGENERAL_CHKRESTORESESSION) )
   gConfig.CheckForUpdates      = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSGENERAL_CHKUPDATES) )


   ' EDITOR OPTIONS   
   HWnd = HWND_FRMOPTIONSEDITOR
   gConfig.SyntaxHighlighting   = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_CHKSYNTAXHIGHLIGHTING) )
   gConfig.LeftMargin           = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_CHKSHOWLEFTMARGIN) )
   gConfig.FoldMargin           = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_CHKSHOWFOLDMARGIN) )
   gConfig.ConfineCaret         = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_CHKCONFINECARET) )
   gConfig.LineNumbering        = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_CHKLINENUMBERING) )
   gConfig.HighlightCurrentLine = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_CHKCURRENTLINE) )
   gConfig.TabIndentSpaces      = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_CHKTABTOSPACES) )
   gConfig.IndentGuides         = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_CHKINDENTGUIDES ) )
   gConfig.RightEdge            = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_CHKSHOWRIGHTEDGE) )
   gConfig.PositionMiddle       = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_CHKPOSITIONMIDDLE) )

   gConfig.RightEdgePosition = AfxGetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_TXTRIGHTEDGE) )
   gConfig.TabSize = AfxGetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR_TXTTABSIZE) )
   gConfig.KeywordCase = ComboBox_GetCurSel( GetDlgItem( HWnd, IDC_FRMOPTIONSEDITOR_COMBOCASE))
   
   ' EDITOR OPTIONS2   
   HWnd = HWND_FRMOPTIONSEDITOR2
   gConfig.Codetips              = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR2_CHKCODETIPS) )
   gConfig.AutoComplete          = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR2_CHKAUTOCOMPLETE) )
   gConfig.CharacterAutoComplete = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR2_CHKCHARAUTOCOMPLETE) )
   gConfig.AutoIndentation       = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR2_CHKAUTOINDENTATION) )
   gConfig.ForNextVariable       = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR2_CHKFORNEXTVARIABLE) )
   gConfig.BraceHighlight        = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR2_CHKBRACEHIGHLIGHT) )
   gConfig.OccurrenceHighlight   = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSEDITOR2_CHKOCCURRENCEHIGHLIGHT) )


   ' If a localization file is currently being edited then save it now.
   frmOptionsLocal_LocalEditCheck

   ' If the localization has changed then display a message indicating that the changes
   ' will not be seen until the application is restarted.
   HWnd = HWND_FRMOPTIONSLOCAL
   wText = AfxGetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSLOCAL_FRAMELOCALIZATION) )
   If (Ucase(wText) <> Ucase(gConfig.LocalizationFile)) orelse (gLocalPhrasesEdit = true) Then
      MessageBox( HWND_FRMMAIN, _
                  L(108,"Localization changes will be applied the next time the application is run."), _
                  L(109,"Warning"), _
                  MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 Or MB_APPLMODAL )
   End If
   gConfig.LocalizationFile = wText
   
   
   ' THEME COLOR/FONT OPTIONS
   HWnd = HWND_FRMOPTIONSCOLORS

   Dim idx As Long 
   idx = ListBox_GetCurSel( GetDlgItem(HWnd, IDC_FRMOPTIONSCOLORS_LSTTHEMES) )
   if idx >= lbound(gThemeFilenames) and idx <= ubound(gThemeFilenames) then
      gConfig.ThemeFilename = gThemeFilenames(idx)
   end if   
   idx = ComboBox_GetCurSel( GetDlgItem(HWnd, IDC_FRMOPTIONSCOLORS_COMBOFONTNAME) )
   gConfig.EditorFontname = AfxGetComboBoxText( GetDlgItem(HWnd, IDC_FRMOPTIONSCOLORS_COMBOFONTNAME), idx )
   idx = ComboBox_GetCurSel( GetDlgItem(HWnd, IDC_FRMOPTIONSCOLORS_COMBOFONTSIZE) )
   gConfig.EditorFontSize = AfxGetComboBoxText( GetDlgItem(HWnd, IDC_FRMOPTIONSCOLORS_COMBOFONTSIZE), idx )  
   idx = ComboBox_GetCurSel( GetDlgItem(HWnd, IDC_FRMOPTIONSCOLORS_COMBOFONTCHARSET) )
   gConfig.EditorFontCharSet = AfxGetComboBoxText( GetDlgItem(HWnd, IDC_FRMOPTIONSCOLORS_COMBOFONTCHARSET), idx )
   gConfig.FontExtraSpace = AfxGetWindowText( GetDlgItem( HWnd, IDC_FRMOPTIONSCOLORS_TXTEXTRASPACE))
   
   
   ' COMPILER OPTIONS
   HWnd = HWND_FRMOPTIONSCOMPILER
   
   ' Get the selected toolchain and then build the paths to the compilers
   dim as CWSTR wszPath
   idx = ListBox_GetCurSel( GetDlgItem(HWnd, IDC_FRMOPTIONSCOMPILER_LSTTOOLCHAINS) )
   if idx <> -1 then
      wszPath = AfxGetExePathName + "Toolchains\" + AfxGetListBoxText( GetDlgItem(HWnd, IDC_FRMOPTIONSCOMPILER_LSTTOOLCHAINS), idx )
      gConfig.FBWINCompiler32 = wszPath + "\fbc32.exe"
      gConfig.FBWINCompiler64 = wszPath + "\fbc64.exe"
   end if
   gConfig.CompilerSwitches = AfxGetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSCOMPILER_TXTFBSWITCHES) )
   gConfig.CompilerHelpfile = AfxGetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSCOMPILER_TXTFBHELPFILE) )
   gConfig.WinFBXHelpfile   = AfxGetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSCOMPILER_TXTWINFBXHELPPATH) )
   gConfig.RunViaCommandWindow = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMOPTIONSCOMPILER_CHKRUNVIACOMMANDWINDOW) )
   gConfig.DisableCompileBeep = Button_GetCheck( GetDlgItem( HWnd, IDC_FRMOPTIONSCOMPILER_CHKDISABLECOMPILEBEEP) )
   
   
   ' KEYWORDS
   If gConfig.bKeywordsDirty Then
      Dim As HWnd hTextBox = GetDlgItem(HWND_FRMOPTIONSKEYWORDS, IDC_FRMOPTIONSKEYWORDS_TXTKEYWORDS)
      gConfig.FBKeywords = AfxGetWindowText(hTextBox)
      hTextBox = GetDlgItem(HWND_FRMOPTIONSKEYWORDSWINAPI, IDC_FRMOPTIONSKEYWORDSWINAPI_TXTKEYWORDS)
      gConfig.WinApiKeywords = AfxGetWindowText(hTextBox)
      gConfig.SaveKeywords
   End If
   
   ' Output the config settings to disk file
   gConfig.SaveConfigFile
   
   ' If Occurrences Highlighting is false then ensure that any highlights on the screen will be removed.
   dim pDoc as clsDocument ptr
   pDoc = gTTabCtl.GetActiveDocumentPtr()
   if pDoc then
      frmMain_HighlightWord( pDoc, "" )
   end if

   function = 0
End function


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmOptions
' ========================================================================================
Function frmOptions_OnCreate( _
            ByVal HWnd As HWnd, _
            ByVal lpCreateStructPtr As LPCREATESTRUCT _
            ) As BOOLEAN

   ' This is a modal popup window so disable the parent window
   DisableAllModeless()

   '  Message cracker macro expects a True to be returned for a successful
   '  OnCreate handler even though returning -1 from a standard WM_CREATE
   '  call would stop creating the window. This is just one of those Windows
   '  inconsistencies.
   Return True
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmOptions
' ========================================================================================
Function frmOptions_OnCommand( _
            ByVal HWnd As HWnd, _
            ByVal id As Long, _
            ByVal hwndCtl As HWnd, _
            ByVal codeNotify As UINT _
            ) As LRESULT

   Select Case id
      
      Case IDC_FRMOPTIONS_CMDOK
         If codeNotify = BN_CLICKED Then
            frmOptions_SaveEditorOptions()
            ' Reload the theme file in case it has changed
            LoadThemeFile()
            ' Refresh the frmOutput controls (listviews and RichEdits)
            AfxRedrawWindow( HWND_FRMOUTPUT )
            frmOutput_SetControlColors()
            frmOutput_PositionWindows()
            ' Refresh the menubar. The rest of the GUI will refresh via frmMain_PositionWindows
            AfxRedrawWindow( HWND_FRMMAIN_MENUBAR )
            frmMenuBar_PositionWindows()
            
            ' Reposition main windows in case hide toolbar/statusbar selected
            frmMain_PositionWindows
            ' Apply the newly saved options to all open Scintilla windows
            dim pFrame as CWindow ptr
            Dim pDoc As clsDocument Ptr = gApp.pDocList
            do until pDoc = 0
               pDoc->ApplyProperties
               ' Ensure the grab handles of form and controls are redrawn or hidden
               if IsWindow(pDoc->hWndFrame) then 
                  pFrame = AfxCWindowPtr(pDoc->hWndFrame)
                  pFrame->Brush = ghDesigner.hPanelBrush
                  AfxRedrawWindow(pDoc->hWndFrame)
               end if   
               pDoc = pDoc->pDocNext
            loop   
            SendMessage( HWnd, WM_CLOSE, 0, 0 )
            Exit Function
         End If
      
      Case IDCANCEL
         If codeNotify = BN_CLICKED Then
            SendMessage( HWnd, WM_CLOSE, 0, 0 )
            Exit Function
         End If
   
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmOptions
' ========================================================================================
Function frmOptions_OnNotify( _
            ByVal HWnd As HWnd, _
            ByVal id As Long, _
            ByVal pNMHDR As NMHDR Ptr _
            ) As LRESULT

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
            dim as long nBranchNum = TreeView_GetlParam( pNMHDR->hWndFrom, lpNMTV->itemNew.hItem)
            Select Case nBranchNum
               Case 0
                  AfxSetWindowText( hLabel, " " & L(226,"General Options")) 
                  hForm = HWND_FRMOPTIONSGENERAL
               Case 1
                  AfxSetWindowText( hLabel, " " & L(110,"Code Editor")) 
                  hForm = HWND_FRMOPTIONSEDITOR
               Case 2
                  AfxSetWindowText( hLabel, " " & L(110,"Code Editor")) 
                  hForm = HWND_FRMOPTIONSEDITOR2
               Case 3
                  AfxSetWindowText( hLabel, " " & L(111,"Themes and Fonts") )
                  hForm = HWND_FRMOPTIONSCOLORS
               Case 4
                  AfxSetWindowText( hLabel, " " & L(112,"Compiler Setup") )
                  hForm = HWND_FRMOPTIONSCOMPILER
               Case 5
                  AfxSetWindowText( hLabel, " " & L(129,"Localization") )
                  hForm = HWND_FRMOPTIONSLOCAL
               Case 6
                  AfxSetWindowText( hLabel, " " & L(177,"FreeBASIC Keywords") )
                  hForm = HWND_FRMOPTIONSKEYWORDS
               Case 7
                  AfxSetWindowText( hLabel, " " & L(435,"Windows Api Keywords") )
                  hForm = HWND_FRMOPTIONSKEYWORDSWINAPI
            End Select
             
            ' Save the most recently opened tree branch 
            OptionsDialogLastOpened = nBranchNum

            
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
Function frmOptions_OnCtlColorStatic( _
            byval HWnd As HWnd, _
            byval hdc As HDC, _
            byval hWndChild As HWnd, _
            byval nType As Long _
            ) As HBRUSH

   if hWndChild = GetDlgItem( HWnd, IDC_FRMOPTIONS_LBLCATEGORY ) then
      ' Set the category label to blue text 
      SetTextColor hdc, 8388608
      SetBkColor   hdc, GetSysColor( COLOR_BTNFACE )
      return GetSysColorBrush( COLOR_BTNFACE )
   end if
   
   function = 0      
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmOptions
' ========================================================================================
Function frmOptions_OnClose( byval HWnd As HWnd ) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableAllModeless()
   DestroyWindow HWnd
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmOptions
' ========================================================================================
Function frmOptions_OnDestroy( byval HWnd As HWnd ) As LRESULT
   Dim As HFONT hFont = AfxGetWindowFont(GetDlgItem(HWnd, IDC_FRMOPTIONS_LBLCATEGORY))
   DeleteFont(hFont)
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' frmOptions Window procedure
' ========================================================================================
Function frmOptions_WndProc( _
            ByVal HWnd   As HWnd, _
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

   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmOptions_Show
' ========================================================================================
public Function frmOptions_Show( ByVal hWndParent As HWnd ) as LRESULT

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
   
   pWindow->AddControl("LABEL", , IDC_FRMOPTIONS_LABEL1, "", 8, 405, 600, 2, _
        WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_NOTIFY Or SS_SUNKEN, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("BUTTON", , IDC_FRMOPTIONS_CMDOK, L(0,"&OK"), 450, 417, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_DEFPUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   
   pWindow->AddControl("BUTTON", , IDCANCEL, L(1,"&Cancel"), 535, 417, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)


   ' Reload the config file in case a user has automated a change to it since
   ' the application started (for example, changing the compiler path).
   gConfig.LoadConfigFile()

   ' Configure the TreeView
   Dim hItems(7) As HTREEITEM 
   hItems(0) = TreeView_AppendItem (hWndTreeview, TVI_ROOT, L(226,"General Options"),    0)
   hItems(1) = TreeView_AppendItem (hWndTreeview, TVI_ROOT, L(110,"Code Editor"),        1)
   hItems(2) = TreeView_AppendItem (hWndTreeview, TVI_ROOT, L(110,"Code Editor"),        2)
   hItems(3) = TreeView_AppendItem (hWndTreeview, TVI_ROOT, L(111,"Colors and Fonts"),   3)
   hItems(4) = TreeView_AppendItem (hWndTreeview, TVI_ROOT, L(112,"Compiler Setup"),     4)
   hItems(5) = TreeView_AppendItem (hWndTreeview, TVI_ROOT, L(129,"Localization"),       5)
   hItems(6) = TreeView_AppendItem (hWndTreeview, TVI_ROOT, L(177,"FreeBASIC Keywords"), 6)
   hItems(7) = TreeView_AppendItem (hWndTreeview, TVI_ROOT, L(435,"Windows API Keywords"), 7)

   ' Load all of the child dialogs
   frmOptionsGeneral_Show( pWindow->hWindow )
   frmOptionsEditor_Show( pWindow->hWindow )
   frmOptionsEditor2_Show( pWindow->hWindow )
   frmOptionsColors_Show( pWindow->hWindow )
   frmOptionsCompiler_Show( pWindow->hWindow )
   frmOptionsLocal_Show( pWindow->hWindow )
   frmOptionsKeywords_Show( pWindow->hWindow )
   frmOptionsKeywordsWinApi_Show( pWindow->hWindow )

   ' Select the most previously open tree branch or first item "Code Editor" by default.
   TreeView_SelectItem( hWndTreeview, hItems(OptionsDialogLastOpened) )

   SetFocus hWndTreeview
   
   ' Process Windows messages
   Function = pWindow->DoEvents(SW_SHOW)
   
   ' Delete the frmOptions CWindow class manually allocated memory 
   Delete pWindow

   ' Delete the child windows releasing their manually allocated memory 
   pWindow = AfxCWindowPtr(HWND_FRMOPTIONSGENERAL):   Delete pWindow
   pWindow = AfxCWindowPtr(HWND_FRMOPTIONSEDITOR):    Delete pWindow
   pWindow = AfxCWindowPtr(HWND_FRMOPTIONSEDITOR2):   Delete pWindow
   pWindow = AfxCWindowPtr(HWND_FRMOPTIONSCOLORS):    Delete pWindow
   pWindow = AfxCWindowPtr(HWND_FRMOPTIONSCOMPILER):  Delete pWindow
   pWindow = AfxCWindowPtr(HWND_FRMOPTIONSLOCAL):     Delete pWindow
   pWindow = AfxCWindowPtr(HWND_FRMOPTIONSKEYWORDS):  Delete pWindow
   pWindow = AfxCWindowPtr(HWND_FRMOPTIONSKEYWORDSWINAPI):  Delete pWindow

End Function





