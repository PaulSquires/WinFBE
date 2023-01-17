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


''  HANDLE RAW INCOMING MESSAGES IN THE MESSAGE PUMP


function handleMouseScrollBar( byval uMsg as MSG ) as boolean
   if (gApp.IsProjectLoading) orelse (gApp.IsFileLoading) then return false
   if HWND_FRMPANEL_VSCROLLBAR = 0 then return false

   ' Determine if the mouse message should be intercepted by the 
   ' transparent WS_EX_LAYERED scrollbar
   select case uMsg.message 
      case WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_LBUTTONUP
         if IsWindowVisible( HWND_FRMPANEL_VSCROLLBAR ) then
            dim as POINT pt: GetCursorPos( @pt )
            dim as RECT rc = AfxGetWindowRect(HWND_FRMPANEL_VSCROLLBAR)
            if PtInRect( @rc, pt ) orelse _
               GetCapture() = HWND_FRMPANEL_VSCROLLBAR then 
               ' forward the message to the scrollbar instead of say 
               ' the Explorer listbox that is underneath it
               AfxRedrawWindow( gPanelVScroll.hListBox )
               SendMessage( HWND_FRMPANEL_VSCROLLBAR, uMsg.message, uMsg.wParam, uMsg.lParam )
               return true
            end if
         end if
   end select
   ' we want our mouse message to continue so return FALSE
   return false
end function


function handleMouseShowScrollBar( byval uMsg as MSG ) as boolean
   if (gApp.IsProjectLoading) orelse (gApp.IsFileLoading) then return false
   
   ' Determine if the mouse has moved over (or away from) a window that should invoke
   ' showing or hiding our custom Panel VScroll or Editor HScroll.
   if (uMsg.message = WM_MOUSEMOVE) orelse (uMsg.message = WM_NCMOUSEMOVE) then
      ' if mouse is currently in drag operation then simply exit
      if gApp.bDragActive then return false
      ' if mouse is currently scrolling through top menus then exit
      if IsWindowVisible( HWND_MENU(0) ) then return false
      if IsWindowVisible( HWND_FRMPANEL ) = 0 then return false
   
      ' HIT TEST for VSCROLLBAR in PANEL
      dim as boolean isVisible = IsWindowVisible(HWND_FRMPANEL_VSCROLLBAR)
      dim as long curState = iif(isVisible, SW_SHOWNA, SW_HIDE)
      dim as long newState = SW_HIDE
      if isMouseOverWindow( gPanelVScroll.hListBox ) orelse _
         isMouseOverWindow( HWND_FRMPANEL_VSCROLLBAR ) then 
         newState = SW_SHOWNA
      end if
      if newState <> curState then frmPanelVScroll_PositionWindows( newState )
   end if      

   ' Should we show an Editor HScrollbar
   dim as clsDocument ptr pDoc = gTTabCtl.GetActiveDocumentPtr()
   if pDoc then
      if (pDoc->IsDesigner) andalso (IsDesignerView(pDoc)) then
         ' visual designer design view is currently active
      else
         for i as long = lbound(pDoc->hWindow) to ubound(pDoc->hWindow)
            dim as boolean isVisible = IsWindowVisible( HWND_FRMEDITOR_HSCROLLBAR(i) )
            dim as long curState = iif(isVisible, SW_SHOWNA, SW_HIDE)
            dim as long newState = SW_HIDE
            if isMouseOverWindow( pDoc->hWindow(i) ) orelse _
               isMouseOverWindow( HWND_FRMEDITOR_HSCROLLBAR(i) ) then 
               newState = frmEditorHScroll_NeedScrollBar( pDoc, i )
            end if
            if newState <> curState then 
               ShowWindow( HWND_FRMEDITOR_HSCROLLBAR(i), newState )
               frmMain_PositionWindows()
            end if
         next
      end if
   end if
   
   ' we want our mouse message to continue so return FALSE
   return false
end function


function handleMouseTopMenu( byval uMsg as MSG ) as boolean
   ' Handle any mouse clicks that are outside of a popup menu in order
   ' to close any open top menu popups.

   if uMsg.message = WM_LBUTTONDOWN then
      if (uMsg.HWnd <> ghWndActiveMenuBarButton) andalso _
         (uMsg.HWnd <> HWND_MENU(0)) andalso _
         (GetParent(uMsg.HWnd) <> HWND_MENU(0)) andalso _
         (uMsg.HWnd <> HWND_MENU(1)) andalso _
         (GetParent(uMsg.HWnd) <> HWND_MENU(1)) then
         killAllPopupMenus()
      end if
   elseif uMsg.message = WM_NCLBUTTONDOWN then
      killAllPopupMenus()
   end if
   ' we want our mouse message to continue so return FALSE
   return false
end function


function handleAltKeyMenuBar( byval uMsg as MSG ) as boolean
   ' 2022-01-08: skip processing the Alt key here because it seems to
   ' interfere with the Alt+Mouse column editing in Scintilla. Need
   ' to investigate a better alternative. The following testing for 
   ' the mouse button does work, but it still results in the menubar
   ' being selected/deselected.
   
'   ' single ALT key pressed, activate the first top menubar button
'   if (uMsg.message = WM_SYSKEYDOWN) and (uMsg.wParam = VK_MENU) then
'      killPopupSubMenus()
'      killPopupMenus()
'      ' if active menubar exists then simply toggle it off and exit
'      dim as HWND hCtrl 
'      if ghWndActiveMenuBarButton then
'         hCtrl = ghWndActiveMenuBarButton
'         ghWndActiveMenuBarButton = 0
'         AfxRedrawWindow(hCtrl)
'         return true
'      else   
'         hCtrl = GetDlgItem(HWND_FRMMAIN_MENUBAR, IDC_MENUBAR_FILE)
'         ghWndActiveMenuBarButton = hCtrl
'         AfxRedrawWindow(hCtrl)
'         return true
'      end if   
'   end if
   ' we want our message to continue so return FALSE
   return false
end function


function handleKeysTopMenu( byval uMsg as MSG ) as boolean
   ' If topmenus or menubar are active then process all keyboard input for them rather than 
   ' passing it off to the system to process.
   if (uMsg.message = WM_KEYDOWN) and (ghWndActiveMenuBarButton <> 0) then
      if uMsg.wParam = VK_RETURN then
         ' simulate pressing ENTER on a menuitem by sending WM_LBUTTONUP message
         ' to the active popup menu or menubar button. If this item hasa submenu 
         ' then convert the key from ENTER to RIGHT ARROW.
         dim as HWND hListBox
         if IsWindowVisible( HWND_MENU(0) ) THEN hListBox = GetDlgItem(HWND_MENU(0), IDC_MENU_LISTBOX)
         if IsWindowVisible( HWND_MENU(1) ) THEN hListBox = GetDlgItem(HWND_MENU(1), IDC_MENU_LISTBOX)
         ' If we have a valid listbox then process the action on the current listbox selection, otherwise
         ' no popup menu exists so simulate pressing enter on the active menubar button.
         if hListBox then
            ' Determine if the menu entry has a child popup. If yes, then convert the keypress to right
            ' arrow and process it below in the VK_RIGHT handler.
            dim as long nCurSel = ListBox_GetCurSel(hListBox)
            dim as long idx = ListBox_GetItemData(hListBox, nCurSel)
            if gTopMenu(idx).nChildID <> 0 then
               uMsg.wParam = VK_RIGHT
            else
               SendMessage( hListBox, WM_LBUTTONUP, 0, 0 )
               return true
            end if   
         else   
            SendMessage( ghWndActiveMenuBarButton, WM_LBUTTONDOWN, 0, 0 )
            return true
         end if
      end if
      
      select case uMsg.wParam
         case VK_ESCAPE
            if IsWindowVisible( HWND_MENU(1) ) THEN 
               killPopupSubMenus()
               return true
            elseif IsWindowVisible( HWND_MENU(0) ) THEN
               killPopupMenus()
               return true
            elseif ghWndActiveMenuBarButton <> 0 then
               dim as HWND hCtrl = ghWndActiveMenuBarButton
               ghWndActiveMenuBarButton = 0
               AfxRedrawWindow(hCtrl)
               return true
            end if   
   
         case VK_LEFT, VK_RIGHT, VK_UP, VK_DOWN
            if IsWindowVisible(HWND_MENU(0)) or IsWindowVisible(HWND_MENU(1)) then   
               if uMsg.wParam = VK_DOWN then 
                  setNextMenuItemTabIndex(false)
                  return true
               end if
               if uMsg.wParam = VK_UP then 
                  setNextMenuItemTabIndex(true)
                  return true
               end if
               if uMsg.wParam = VK_LEFT then 
                  if IsWindowVisible(HWND_MENU(1)) then
                     killPopupSubMenus()
                     return true
                  end if   
               end if
               if uMsg.wParam = VK_RIGHT then
                  if IsWindowVisible(HWND_MENU(0)) then
                     dim as HWND hListBox = GetDlgItem(HWND_MENU(0), IDC_MENU_LISTBOX)
                     gMenuLastCurSel = ListBox_GetCurSel(hListBox)
                     if frmPopupMenu_Show( ID_SUBPOPUP, gMenuLastCurSel, hListBox ) then
                        setNextMenuItemTabIndex(false)
                        return true
                     end if   
                  end if
               end if
            end if
      
            if ghWndActiveMenuBarButton <> 0 then
               dim as HWND prevCtrl = ghWndActiveMenuBarButton
               if uMsg.wParam = VK_LEFT then 
                  ghWndActiveMenuBarButton = GetNextDlgGroupItem( HWND_FRMMAIN_MENUBAR, ghWndActiveMenuBarButton, true )
                 ' if popup menus are already activated then continue to show them as
                 ' we move through the various menubar buttons.
                  if IsWindowVisible(HWND_MENU(0)) then
                     frmPopupMenu_Show(ID_POPUP, 0, ghWndActiveMenuBarButton)
                     setNextMenuItemTabIndex(false)
                  end if   
               end if
               if uMsg.wParam = VK_RIGHT then 
                  ghWndActiveMenuBarButton = GetNextDlgGroupItem( HWND_FRMMAIN_MENUBAR, ghWndActiveMenuBarButton, false )
                  ' if popup menus are already activated then continue to show them as
                  ' we move through the various menubar buttons.
                  if IsWindowVisible(HWND_MENU(0)) then
                     frmPopupMenu_Show(ID_POPUP, 0, ghWndActiveMenuBarButton)
                     setNextMenuItemTabIndex(false)
                  end if   
               end if
               if uMsg.wParam = VK_DOWN then 
                  frmPopupMenu_Show(ID_POPUP, 0, ghWndActiveMenuBarButton)
                  setNextMenuItemTabIndex(false)
               end if   
               AfxRedrawWindow(prevCtrl)
               AfxRedrawWindow(ghWndActiveMenuBarButton)
               return true
            end if
      
         case else
            return true
      end select
   end if

   ' we want our message to continue so return FALSE
   return false
end function


function handleKeysExplorerListBox( byval uMsg as MSG ) as boolean
   ' Was ENTER key pressed while scrolling the Explorer listbox via keyboard
   dim as HWND hList = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_LISTBOX)
   if (uMsg.message = WM_KEYDOWN) andalso (uMsg.HWnd = hList) then
      dim as long nSel = ListBox_GetCaretIndex( hList )
      dim as CWSTR wszCaption = AfxGetListBoxText( hList, nSel )
      if uMsg.wParam = VK_RETURN then
         if left(wszCaption, 1) = "%" then
            ' expand/contract the header node
            dim as long idxArray = getExplorerNodeHeaderIndex( wszCaption )
            if idxArray <> -1 then
               gConfig.Cat(idxArray).bShow = not gConfig.Cat(idxArray).bShow
            end if
            LoadExplorerFiles()
            SetFocus( hList )
            return true
         else   
            ' open the selected document
            dim as CWSTR wszFilename
            dim as long p = ListBox_GetItemData( hList, nSel )
            dim pDoc as clsDocument ptr = cast( clsDocument ptr, cast(integer, p) )
            if pDoc then
               wszFilename = pDoc->DiskFilename
               OpenSelectedDocument( wszFilename, "" )
            end if
            SetFocus( hList )
            return true
         end if
      end if
   end if

   ' we want our message to continue so return FALSE
   return false
end function


function handleEscKeyModeless( byval uMsg as MSG ) as boolean
   ' Close any open Find/Replace window with ESC key
   ' Close any open FindInFiles window with ESC key
   ' Close any open FunctionList window
   if (uMsg.message = WM_KEYDOWN) andAlso (uMsg.wParam = VK_ESCAPE) THEN
      if IsWindowVisible( HWND_FRMFINDREPLACE ) THEN 
         SendMessage( HWND_FRMFINDREPLACE, WM_CLOSE, 0, 0 )
         return true
      end if
      if IsWindowVisible( HWND_FRMFINDINFILES ) THEN 
         SendMessage( HWND_FRMFINDINFILES, WM_CLOSE, 0, 0 )
         return true
      end if
   end if
   ' we want our message to continue so return FALSE
   return false
end function


function handleKeysFindReplace( byval uMsg as MSG ) as boolean
   ' For Find textbox, an ENTER will move to the next selection. A SHIFT+ENTER
   ' will move to the previous. For Replace textbox, use the same logic but 
   ' also perform the replace action.
   dim as HWND hFind = GetDlgItem( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND )
   dim as HWND hReplace = GetDlgItem( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTREPLACE )
   if (GetFocus = hFind) orelse (GetFocus = hReplace) then
      if (uMsg.message = WM_KEYDOWN) and (uMsg.wParam = VK_RETURN) then
         if (GetAsyncKeyState(VK_CONTROL) and &H8000) andalso _
            (GetAsyncKeyState(VK_MENU) and &H8000) THEN
            ' Ctrl+Alt+Enter  Replace All
            SendMessage( HWND_FRMMAIN, WM_COMMAND, MAKEWPARAM(IDM_REPLACEALL, 0), 0 )
            return true
         end if
         if (GetAsyncKeyState(VK_SHIFT) and &H8000) THEN
            if GetFocus = hFind then
               SendMessage( HWND_FRMMAIN, WM_COMMAND, MAKEWPARAM(IDM_FINDPREV, 0), 0 )
            elseif GetFocus = hReplace then
               SendMessage( HWND_FRMMAIN, WM_COMMAND, MAKEWPARAM(IDM_REPLACEPREV, 0), 0 )
            end if   
            return true
         else   
            if GetFocus = hFind then
               SendMessage( HWND_FRMMAIN, WM_COMMAND, MAKEWPARAM(IDM_FINDNEXT, 0), 0 )
            elseif GetFocus = hReplace then
               SendMessage( HWND_FRMMAIN, WM_COMMAND, MAKEWPARAM(IDM_REPLACENEXT, 0), 0 )
            end if
            return true
         end if   
      end if
   end if
   ' we want our message to continue so return FALSE
   return false
end function


function handleKeysVisualDesigner( byval uMsg as MSG ) as boolean
  ' Handle any keypress that would move or resize control(s) on a Designer Form.
   if uMsg.message = WM_KEYDOWN THEN
      dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
      if pDoc THEN
         ' We don't want keystrokes going into the scintilla control if the
         ' visual designer form is active.
         if pDoc->IsDesigner THEN
            if IsDesignerView(pDoc) then
               ' Don't allow keys to the scintilla control
               if pDoc->IsValidScintillaID(GetDlgCtrlID(GetFocus)) then 
                  if uMsg.wParam = VK_DELETE then
                     frmMain_OnCommand(HWND_FRMMAIN, IDM_DELETE, 0, 0)
                     return true
                  end if   
               end if
            end if
         end if
      end if         

      select CASE uMsg.wParam
         case VK_SPACE
            dim as HWND hEventList = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTEVENTS)
            if pDoc andalso GetFocus() = hEventList then
               ' Toggle the selected Event listbox item
               dim as long nCurSel = ListBox_GetCurSel(hEventList)
               if nCurSel > -1 then 
                  dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
                  dim pCtrl as clsControl ptr 
                  if pDoc then
                     pCtrl = pDoc->Controls.GetActiveControl
                     if pCtrl THEN
                        pCtrl->Events(nCurSel).bIsSelected = not pCtrl->Events(nCurSel).bIsSelected
                        AfxRedrawWindow(hEventList)
                        pDoc->UserModified = true
                        pDoc->bRegenerateCode = true
                        frmMain_SetStatusbar
                        return true
                     end if
                  end if
               end if   
            end if
                        
         case VK_UP, VK_DOWN, VK_LEFT, VK_RIGHT
            if pDoc THEN
               if pDoc->IsDesigner THEN
                  select case GetFocus
                     ' Allow arrows to move selected item in a listbox or edit textbox
                     case GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTTOOLBOX), _
                          GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTPROPERTIES), _
                          GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTEVENTS), _
                          HWND_PROPLIST_EDIT
                     case else
                        ' Allow arrows to move/resize selected controls
                        if IsDesignerView(pDoc) then
                           if pDoc->Controls.SelectedControlsCount > 0 then
                              ' Move control(s)
                              if (GetAsyncKeyState(VK_CONTROL) and &H8000) then
                                 KeyboardMoveControls(pDoc, uMsg.wParam)
                                 return true
                              end if
                              ' Resize control(s)   
                              if (GetAsyncKeyState(VK_SHIFT) and &H8000) THEN
                                 KeyboardResizeControls(pDoc, uMsg.wParam)
                                 return true
                              end if   
                              ' Move active selected control focus within group of controls
                              KeyboardCycleActiveControls(pDoc, uMsg.wParam)
                              return true
                           end if   
                        end if   
                  end select
               end if
            end if
      end select   
   end if  

   ' we want our message to continue so return FALSE
   return false
end function

      