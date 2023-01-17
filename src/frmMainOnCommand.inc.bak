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

#include once "frmMainOnCommand.bi"

#include once "frmOutput.bi"
#include once "clsApp.bi"
#include once "clsDocument.bi"
#include once "frmMain.bi"
#include once "frmFunctions.bi"


' Global array to hold cut/copy/paste controls
dim shared gCopyControls(any) as clsControl


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmMain
' ========================================================================================
Function frmMain_OnCommand( _
            ByVal HWnd As HWnd, _
            ByVal id As Long, _
            ByVal hwndCtl As HWnd, _
            ByVal codeNotify As UINT _
            ) As LRESULT
   Dim As HWnd hEdit 

   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   If pDoc Then 
      select case codeNotify
         case SCEN_SETFOCUS
            pDoc->hWndActiveScintilla = hwndCtl

            ' Remove any markers set in the document that highlights
            ' entire lines (used for Selected text searching).
            pDoc->RemoveMarkerHighlight

      end select      
      hEdit = pDoc->hWndActiveScintilla
   End If   

   ' If the popup Help Viewer window has focus then exit function because we don't
   ' want accelerator keys to act on the code editor.
   if IsWindow( HWND_FRMHELPVIEWER ) then
      if GetForegroundWindow() = HWND_FRMHELPVIEWER then 
         if id = IDM_FIND then
            frmHelpViewer_OnCommand(HWND_FRMHELPVIEWER, IDC_FRMHELPVIEWER_FIND, 0, BN_CLICKED)
         end if   
         exit function
      end if
   end if
   
   ' Kill any currently displayed topmenu whenever a popup dialog is about to be activated
   Select Case id
      case IDM_FILEOPEN, IDM_FILEOPENTEMPLATES, IDM_FILESAVEAS, _
           IDM_SAVESESSION, IDM_LOADSESSION, IDM_CATEGORIES, _
           IDM_KEYBOARDSHORTCUTS, IDM_USERTOOLSDIALOG, IDM_OPTIONSDIALOG, IDM_BUILDCONFIG, _
           IDM_USERSNIPPETS, IDM_FIND, IDM_FINDINFILES, IDM_PROJECTOPEN, IDM_PROJECTSAVEAS, _
           IDM_PROJECTOPTIONS, IDM_COMMANDLINE, IDM_ABOUT
         killAllPopupMenus()
   end select
   
   Select Case id
       ' show any topmenu that were activated via keyboard Alt+ sequences     
      case IDC_MENUBAR_FILE to IDC_MENUBAR_HELP
         killAllPopupMenus()
         dim as HWND hCtrl = GetDlgItem(HWND_FRMMAIN_MENUBAR, id)
         ghWndActiveMenuBarButton = hCtrl
         SendMessage( hCtrl, WM_LBUTTONDOWN, 0, 0)     
         setNextMenuItemTabIndex(false)

      case IDM_PROJECTFILETYPE
         dim as HMENU hPopupMenu = CreateStatusBarFileTypeContextMenu()
         dim as POINT pt: GetCursorPos @pt
         TrackPopupMenu(hPopUpMenu, 0, pt.x, pt.y, 0, HWnd, ByVal Null)
         DestroyMenu hPopUpMenu

      case IDM_FILEENCODING
         dim as HMENU hPopupMenu = CreateStatusBarFileEncodingContextMenu()
         dim as POINT pt: GetCursorPos @pt
         TrackPopupMenu(hPopUpMenu, TPM_RIGHTALIGN, pt.x, pt.y, 0, HWnd, ByVal Null)
         DestroyMenu hPopUpMenu
        
      case IDM_LINEENDINGS
         dim as HMENU hPopupMenu = CreateStatusBarLineEndingsContextMenu()
         dim as POINT pt: GetCursorPos @pt
         TrackPopupMenu(hPopUpMenu, TPM_RIGHTALIGN, pt.x, pt.y, 0, HWnd, ByVal Null)
         DestroyMenu hPopUpMenu

      case IDM_SPACES
         dim as HMENU hPopupMenu = CreateStatusBarSpacesContextMenu()
         dim as POINT pt: GetCursorPos @pt
         dim as long idMenu = TrackPopupMenu(hPopUpMenu, TPM_RETURNCMD, pt.x, pt.y, 0, HWnd, ByVal Null)
         DestroyMenu hPopUpMenu
         if (idMenu >= IDM_SPACES) andalso (idMenu <= IDM_SPACES + 8) then
            gConfig.TabSize = idMenu - IDM_SPACES
            gSBPanels(4).wszText = "Spaces: " & gConfig.TabSize
            frmStatusBar_PositionWindows
            Dim pDoc As clsDocument Ptr = gApp.pDocList
            do until pDoc = 0
               pDoc->ApplyProperties
               pDoc = pDoc->pDocNext
            loop   
         end if
   
      case IDM_CLOSEPANEL
         ShowWindow( HWND_FRMPANEL, SW_HIDE )
         frmMain_PositionWindows()
         
      case IDM_OPENINCLUDE:           OpenSelectedDocument( gApp.IncludeFilename )
      case IDM_SETFOCUSEDITOR:        frmMain_SetFocusToCurrentCodeWindow()   
      case IDM_EXPLORER_EXPANDALL:    frmExplorer_ExpandAll()   
      case IDM_EXPLORER_COLLAPSEALL:  frmExplorer_CollapseAll()   
      case IDM_FUNCTIONS_EXPANDALL:   frmFunctions_ExpandAll()   
      case IDM_FUNCTIONS_COLLAPSEALL: frmFunctions_CollapseAll()   
      case IDM_BOOKMARKS_EXPANDALL:   frmBookmarks_ExpandAll()   
      case IDM_BOOKMARKS_COLLAPSEALL: frmBookmarks_CollapseAll()   
      case IDM_FUNCTIONS_VIEWASTREE:  frmFunctions_ViewAsTree()
      case IDM_FUNCTIONS_VIEWASLIST:  frmFunctions_ViewAsList()


      ''  FILE MENU
      Case IDM_FILENEW:            OnCommand_FileNew( HWnd )
      Case IDM_FILEOPEN:           OnCommand_FileOpen( HWnd, true )
      Case IDM_FILEOPENTEMPLATES:  OnCommand_FileTemplates( HWnd )
      Case IDM_FILESAVE:           OnCommand_FileSave( HWnd, pDoc, false, false )
      Case IDM_FILESAVEAS:         OnCommand_FileSave( HWnd, pDoc, true, false )
      Case IDM_FILECLOSE:          OnCommand_FileClose( HWnd, EFC_CLOSECURRENT )
      Case IDM_FILESAVEALL:        OnCommand_FileSaveAll( HWnd ) 
      case IDM_AUTOSAVE:           OnCommand_FileAutoSave( HWnd )
      case IDM_LOADSESSION:        OnCommand_FileLoadSession( HWnd )   
      case IDM_SAVESESSION:        OnCommand_FileSaveSession( HWnd )   
      Case IDM_FILECLOSEALL:       OnCommand_FileClose( HWnd, EFC_CLOSEALL )
      Case IDM_FILECLOSEALLOTHERS: OnCommand_FileClose( HWnd, EFC_CLOSEALLOTHERS )
      Case IDM_CLOSEALLFORWARD:    OnCommand_FileClose( HWnd, EFC_CLOSEALLFORWARD )
      Case IDM_CLOSEALLBACKWARD:   OnCommand_FileClose( HWnd, EFC_CLOSEALLBACKWARD )
      Case IDM_KEYBOARDSHORTCUTS:  frmKeyboard_Show( HWnd )
      Case IDM_OPTIONSDIALOG:      frmOptions_Show( HWnd )
      case IDM_BUILDCONFIG:        frmBuildConfig_Show( HWND ) 
      case IDM_USERTOOLSDIALOG:    frmUserTools_Show( HWND ) 
      case IDM_USERSNIPPETS:       frmSnippets_Show( HWND ) 
      case IDM_CATEGORIES:         frmCategories_Show( HWND ) 
      Case IDM_MRUCLEAR:           ClearMRUlist( id )

      Case IDM_MRUBASE To (IDM_MRUBASE + 10)
         OpenMRUFile( HWnd, id )

      Case IDM_USERTOOLSBASE To (IDM_USERTOOLSBASE + UBound(gConfig.Tools))
         frmUserTools_ExecuteUserTool( id - IDM_USERTOOLSBASE )

      Case IDM_EXIT
         SendMessage( HWnd, WM_CLOSE, 0, 0 )


      ''  EDIT MENU
      Case IDM_CUT, WM_CUT:        OnCommand_EditCut( pDoc, hEdit )
      Case IDM_COPY, WM_COPY:      OnCommand_EditCopy( pDoc, hEdit )
      Case IDM_PASTE, WM_PASTE:    OnCommand_EditPaste( pDoc, hEdit )
      Case IDM_SELECTALL:          OnCommand_EditSelectAll( pDoc, hEdit )
      Case IDM_FIND:               OnCommand_EditFindDialog()
      case IDM_REPLACE:            OnCommand_EditReplaceDialog()
      Case IDM_FINDINFILES:        OnCommand_EditFindInFiles( hEdit )
      case IDM_INDENTBLOCK:        OnCommand_EditIndentBlock( pDoc, hEdit )
      case IDM_UNINDENTBLOCK:      OnCommand_EditUnIndentBlock( pDoc, hEdit )

      Case IDM_FINDNEXT, IDM_FINDPREV, _
           IDM_FINDNEXTACCEL, IDM_FINDPREVACCEL   ' F3/Shift+F3
           OnCommand_EditFindActions( id, pDoc )
           
      Case IDM_REPLACENEXT, IDM_REPLACEPREV, IDM_REPLACEALL
           OnCommand_EditReplaceActions( id, pDoc )

      case IDM_ANSI, IDM_UTF8BOM, IDM_UTF16BOM
           OnCommand_EditEncoding( id, pDoc )

      Case IDM_DELETELINE,     IDM_INSERTFILE,     IDM_COMMENTBLOCK, _
           IDM_UNCOMMENTBLOCK, IDM_DUPLICATELINE,  IDM_MOVELINEUP, _
           IDM_MOVELINEDOWN,   IDM_TOUPPERCASE,    IDM_NEWLINEBELOWCURRENT, _
           IDM_TOLOWERCASE,    IDM_TOMIXEDCASE,    IDM_EOLTOCRLF, _
           IDM_EOLTOCR,        IDM_EOLTOLF,        IDM_TABSTOSPACES, _
           IDM_SELECTLINE,     IDM_REDO,           IDM_UNDO, _
           WM_UNDO
           OnCommand_EditCommon( id, pDoc )


      ''  SEARCH MENU              
      Case IDM_DEFINITION:           OnCommand_SearchGotoDefinition( pDoc )
      Case IDM_LASTPOSITION:         OnCommand_SearchGotoLastPosition()
      case IDM_GOTONEXTCOMPILEERROR: OnCommand_SearchGotoCompileError( true )
      case IDM_GOTOPREVCOMPILEERROR: OnCommand_SearchGotoCompileError( false )

      case IDM_GOTOHEADERFILE,   IDM_GOTOSOURCEFILE, IDM_GOTOMAINFILE, _ 
           IDM_GOTORESOURCEFILE, IDM_GOTO,           IDM_GOTONEXTFUNCTION, _
           IDM_GOTOPREVFUNCTION, IDM_GOTONEXTTAB,    IDM_GOTOPREVTAB, _
           IDM_CLOSETAB
         OnCommand_SearchGotoFile( id, pDoc )
         
      Case IDM_BOOKMARKTOGGLE,   IDM_BOOKMARKNEXT,   IDM_BOOKMARKPREV, _
           IDM_BOOKMARKCLEARALL, IDM_BOOKMARKCLEARALLDOCS
           OnCommand_SearchBookmarks( id, pDoc )
      
      
      ''  VIEW MENU
      Case IDM_VIEWEXPLORER:       OnCommand_ViewExplorer():
      Case IDM_VIEWOUTPUT:         OnCommand_ViewOutput():
      Case IDM_FUNCTIONLIST:       OnCommand_ViewFunctionList():
      Case IDM_BOOKMARKSLIST:      OnCommand_ViewBookmarksList():
      case IDM_VIEWTODO:           OnCommand_ViewToDo():
      case IDM_VIEWNOTES:          OnCommand_ViewNotes():
      Case IDM_RESTOREMAIN:        OnCommand_ViewRestoreMain():
         
      Case IDM_FOLDTOGGLE, IDM_FOLDBELOW, IDM_FOLDALL, IDM_UNFOLDALL
         OnCommand_ViewFold( id, pDoc )

      Case IDM_ZOOMIN, IDM_ZOOMOUT
         OnCommand_ViewZoom( id, pDoc )

      
      ''  PROJECT MENU
      Case IDM_PROJECTNEW:            OnCommand_ProjectNew( HWnd )
      Case IDM_PROJECTSAVE:           OnCommand_ProjectSave( HWnd, False )
      Case IDM_PROJECTSAVEAS:         OnCommand_ProjectSave( HWnd, True )
      Case IDM_PROJECTOPEN:           OnCommand_ProjectOpen( HWnd )
      Case IDM_PROJECTFILESADD:       OnCommand_ProjectFilesAdd( HWnd )
      Case IDM_PROJECTCLOSE:          OnCommand_ProjectClose( HWnd )
      Case IDM_PROJECTOPTIONS:        frmProjectOptions_Show( HWnd, false ) 
      Case IDM_MRUPROJECTCLEAR:       ClearMRUlist( id )
      Case IDM_REMOVEFILEFROMPROJECT: OnCommand_ProjectRemove( id, pDoc )
               
      Case IDM_SETFILENORMAL, IDM_SETFILEMODULE, IDM_SETFILEMAIN, _
           IDM_SETFILERESOURCE, IDM_SETFILEHEADER
           OnCommand_ProjectSetFileType( id, pDoc )
           
      Case IDM_MRUPROJECTBASE To (IDM_MRUPROJECTBASE + 10)
           OpenMRUProjectFile( HWnd, id )

           
      ''  COMPILE MENU          
      Case IDM_BUILDEXECUTE, IDM_COMPILE, IDM_REBUILDALL, IDM_QUICKRUN, _
           IDM_RUNEXE, IDM_COMMANDLINE
         OnCommand_CompileCommon( id )


      ''   DESIGNER MENU
      case IDM_NEWFORM:            OnCommand_DesignerNewForm( HWnd )
      case IDM_DELETE:             OnCommand_DesignerDeleteKey( pDoc )
      case IDM_VIEWTOOLBOX:        OnCommand_DesignerToolBox( HWnd )
      case IDM_MENUEDITOR:         OnCommand_DesignerMenuEditor( HWnd, pDoc )  
      case IDM_TOOLBAREDITOR:      OnCommand_DesignerToolBarEditor( HWnd, pDoc )       
      case IDM_STATUSBAREDITOR:    OnCommand_DesignerStatusBarEditor( HWnd, pDoc, codeNotify )                 
      case IDM_IMAGEMANAGER:       OnCommand_DesignerImageManager( HWnd, pDoc )                 
      case IDM_SNAPLINES:          OnCommand_DesignerSnapLines( pDoc )                 
      case IDM_LOCKCONTROLS:       OnCommand_DesignerLockControls( pDoc )
         
      case IDM_ALIGNLEFTS, IDM_ALIGNCENTERS, IDM_ALIGNRIGHTS, _
           IDM_ALIGNTOPS, IDM_ALIGNMIDDLES, IDM_ALIGNBOTTOMS, _
           IDM_SAMEWIDTHS, IDM_SAMEHEIGHTS, IDM_SAMEBOTH            
         OnCommand_DesignerAlign( HWND, id )

      case IDM_CENTERHORIZ, IDM_CENTERVERT, IDM_CENTERBOTH
         OnCommand_DesignerCenter( HWND, id )
         
      case IDM_HORIZEQUAL, IDM_HORIZINCREASE, IDM_HORIZDECREASE, IDM_HORIZREMOVE
         OnCommand_DesignerHorizSpacing( hwnd, id )

      case IDM_VERTEQUAL, IDM_VERTINCREASE, IDM_VERTDECREASE, IDM_VERTREMOVE
         OnCommand_DesignerVertSpacing( hwnd, id )
      

      ''   HELP MENU
      Case IDM_HELP, IDM_HELPSHORTCUTS, IDM_HELPWINFBE, IDM_HELPWINFBX
         ShowContextHelp(id) 
  
      case IDM_CHECKFORUPDATES
         DisableAllModeless()
         DoCheckForUpdates( hwnd, false )
         EnableAllModeless()
         
      Case IDM_ABOUT
         DisableAllModeless()
         frmAbout_Show( hwnd )
         EnableAllModeless()
         frmMain_SetFocusToCurrentCodeWindow()
        
   End Select

   Function = 0
End Function


