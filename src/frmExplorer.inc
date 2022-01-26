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
  

''
''  frmExplorer.inc
''   

#include once "frmExplorer.bi"
#include once "clsDB2.bi"

' ========================================================================================
' Unselect all items in the multiselect listbox and reset it to the currently
' active top tab (if any)
' ========================================================================================
function frmExplorer_UnSelectListBox() as long
   for i as long = 0 to ListBox_GetCount( HWND_FRMEXPLORER_LISTBOX ) - 1 
      ListBox_SetSel( HWND_FRMEXPLORER_LISTBOX, false, i )
   next 
   ' force a listbox selection for the currently active tab                   
   dim as long idx = gTTabCtl.CurSel: gTTabCtl.CurSel = -1
   gTTabCtl.SetFocusTab(idx)
   function = 0
end function

' ========================================================================================
' Select the listbox item that matches the incoming pDoc item. Do not open nodes to
' find a possible hidden document because the user may have purposely closed a node
' and we should not re-open it automatically.
' ========================================================================================
function frmExplorer_SelectItemData( byval pDoc as clsDocument ptr ) as boolean
   ' Select the Explorer listbox item where the ItemData holds the pDoc handle.
   ' Reset any other single or multiple selections. Process entire listbox. 
   dim as long idx = -1
   for i as long = 0 to ListBox_GetCount( HWND_FRMEXPLORER_LISTBOX ) - 1 
      ListBox_SetSel( HWND_FRMEXPLORER_LISTBOX, false, i )
      if ListBox_GetItemData(HWND_FRMEXPLORER_LISTBOX, i) = pDoc then
         idx = i
      end if   
   next 

   ListBox_SetSel( HWND_FRMEXPLORER_LISTBOX, true, idx )
   AfxRedrawWindow( HWND_FRMEXPLORER_LISTBOX )
   
   if idx >= 0 then function = true   
end function


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
function frmExplorer_PositionWindows() as LRESULT

   dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMEXPLORER)
   if pWindow = 0 then exit function
   
   ' Get the entire client area
   dim as Rect rc
   GetClientRect(HWND_FRMEXPLORER, @rc)

   ' The Explorer label was painted in the top area so start our listbox immediately after that

   ' show/hide of the listbox is done separately in order to prevent blank client area from 
   ' painting before the lists are loaded. Allow a 2 pixel right border between listbox and
   ' form edge in order to allow for panel resizing. 
   
   dim as long nLeft = 0
   dim as long nTop = rc.top + pWindow->ScaleY(46)
   dim as long nWidth = rc.right  - rc.Left - pWindow->ScaleX(2)
   dim as long nHeight = rc.bottom - rc.top - nTop

   SetWindowPos( HWND_FRMEXPLORER_LISTBOX, 0, nLeft, nTop, nWidth, nHeight, SWP_NOZORDER )

   function = 0
end function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmExplorer
' ========================================================================================
function frmExplorer_OnSize( _
            byval HWnd as HWnd, _
            byval state as UINT, _
            byval cx as long, _
            byval cy as long _
            ) as LRESULT
   if state <> SIZE_MINIMIZED then
      ' Position all of the child windows
      frmExplorer_PositionWindows
   end if
   function = 0
end function
       

' ========================================================================================
' Process WM_PAINT message for window/dialog: frmExplorer
' ========================================================================================
function frmExplorer_OnPaint( byval HWnd as HWnd ) as LRESULT
            
   dim as PAINTSTRUCT ps
   dim as HDC hDc
   
   hDC = BeginPaint(hWnd, @ps)
   FillRect( hDC, @ps.rcPaint, ghExplorer.hPanelBrush )
   EndPaint( hWnd, @ps )
   
   function = 0
end function


' ========================================================================================
' Process WM_MOUSEMOVE message for window/dialog: frmExplorer
' ========================================================================================
function frmExplorer_OnMouseMove( _
            byval HWnd as HWnd, _
            byval x as long, _
            byval y as long, _
            byval keyflags as UINT _
            ) as long

   dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMMAIN)
   if pWindow = 0 then exit function

   ' These are relative to <hWnd>
   dim as POINT pt = (x, y)
  
   ' PANELS LEFT/RIGHT SPLITTER
   dim as Rect rc
   GetWindowRect HWND_FRMEXPLORER, @rc
   rc.Left = rc.Right - pWindow->ScaleX(3)
   GetCursorPos(@pt)
   if PtInRect( @rc, pt ) then
      if WindowFromPoint(pt) = HWND_FRMEXPLORER then
         SetCursor LoadImage( null, MAKEINTRESOURCEW(OCR_SIZEWE), _
                                IMAGE_CURSOR, 0, 0, LR_DEFAULTSIZE or LR_SHARED )
      end if
   end if

   if gApp.bDragActive then
      if gApp.hWndPanel = HWND_FRMEXPLORER then
         ' Get the current rect of the frmExplorer and compare right side to the current cursor position
         ' and then move the right side to equal the cursor position.
         dim as long nDiff = pt.x - rc.Right
         GetClientRect HWND_FRMEXPLORER, @rc
         dim as long nWidth = (rc.Right - rc.Left + nDiff)
         ' make sure width does not go below a specific size otherwise the user will not
         ' be able to grab the panel edge to resize
         nWidth = max( nWidth, pWindow->ScaleX(80) )
         SetWindowPos HWND_FRMEXPLORER, 0, 0, 0, nWidth, rc.Bottom - rc.Top, SWP_NOMOVE or SWP_NOZORDER
         frmMain_PositionWindows
         AfxDoEvents   ' allow screen to repaint
         exit function
      end if
   end if
      
   function = 0
end function


' ========================================================================================
' Process WM_LBUTTONDOWN message for window/dialog: frmExplorer
' ========================================================================================
function frmExplorer_OnLButtonDown( _
            byval HWnd as HWnd, _
            byval fDoubleClick as boolean, _
            byval x as long, _
            byval y as long, _
            byval keyflags as UINT _
            ) as long

   dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMMAIN)
   if pWindow = 0 then exit function

   ' PANELS LEFT/RIGHT SPLITTER
   dim as Rect rc
   dim as Point pt 

   gApp.bDragActive = false 
   GetWindowRect HWND_FRMEXPLORER, @rc
   rc.Left = rc.Right - pWindow->ScaleX(3)
   GetCursorPos(@pt)
   if PtInRect( @rc, pt ) then
      if WindowFromPoint(pt) = HWND_FRMEXPLORER then
         SetCursor LoadImage( null, MAKEINTRESOURCEW(OCR_SIZEWE), _
                                IMAGE_CURSOR, 0, 0, LR_DEFAULTSIZE or LR_SHARED )
         gApp.bDragActive = true 
         gApp.hWndPanel   = HWND_FRMEXPLORER
         SetCapture( HWND_FRMEXPLORER )
      end if
   end if

   function = 0
end function


' ========================================================================================
' Process WM_MEASUREITEM message for window/dialog: frmExplorer
' ========================================================================================
function frmExplorer_OnMeasureItem( _
            byval HWnd as HWnd, _
            byval lpmis as MEASUREITEMSTRUCT ptr _
            ) as long
   ' Set the height of the list box items. 
   dim pWindow as CWindow ptr = AfxCWindowPtr(HWnd)
   lpmis->itemHeight = pWindow->ScaleY(EXPLORERITEM_HEIGHT)
      
   function = 0
end function


' ========================================================================================
' Process WM_DRAWITEM message for window/dialog: frmExplorer
' ========================================================================================
function frmExplorer_OnDrawItem( _
            byval HWnd as HWnd, _
            byval lpdis as const DRAWITEMSTRUCT ptr _
            ) as long

   dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMMAIN)
   if pWindow = 0 then exit function

   if lpdis = 0 then exit function
 
   if ( lpdis->itemAction = ODA_DRAWENTIRE ) orelse _
      ( lpdis->itemAction = ODA_SELECT ) orelse _
      ( lpdis->itemAction = ODA_FOCUS ) then

      if lpdis->itemAction = ODA_FOCUS then return true

      dim as RECT rc = lpdis->rcItem
      dim as long nWidth  = rc.right-rc.left
      dim as long nHeight = rc.bottom-rc.top 

      SaveDC(lpdis->hDC)

      dim memDC as HDC      ' Double buffering
      dim hbit as HBITMAP   ' Double buffering

      memDC = CreateCompatibleDC( lpdis->hDC )
      hbit  = CreateCompatibleBitmap( lpdis->hDC, nWidth, nHeight )
      if hbit then hbit = SelectObject( memDC, hbit )

      SelectObject( memDC, ghMenuBar.hFontMenuBar )
         
      ' Default to using normal
      dim as HBRUSH hBrush = ghExplorer.hBackBrush
      dim as COLORREF foreclr = ghExplorer.ForeColor
      dim as COLORREF backclr = ghExplorer.BackColor
         
      dim as boolean IsHot = false
      dim as boolean isNodeHeader = false
      dim as boolean isIconDown = false
         
      dim as POINT pt
      GetCursorPos( @pt )
      MapWindowPoints( lpdis->hwndItem, HWND_DESKTOP, cast( POINT ptr, @rc ), 2 )
      if PtInRect( @rc, pt ) then IsHot = true
      
      ' if mouse is over VScrollBar then reset hot
      if isMouseOverWindow( HWND_FRMEXPLORER_VSCROLLBAR ) then IsHot = false
         
      if ListBox_GetSel(lpdis->hwndItem, lpdis->itemID) then IsHot = true

      hBrush = iif( IsHot, ghExplorer.hBackBrushHot, ghExplorer.hBackBrush)
      backclr = iif( IsHot, ghExplorer.BackColorHot, ghExplorer.BackColor)
      foreclr = iif( IsHot, ghExplorer.ForeColorHot, ghExplorer.ForeColor)

      dim as CWSTR wszCaption = AfxGetListBoxText(lpdis->hwndItem, lpdis->ItemID)

      ' if this is a "node" header then use those colors
      if left(wszCaption, 1) = "%" then isNodeHeader = true
      if mid(wszCaption, 4, 1) = "1" then isIconDown = true
      
      if isNodeHeader then
         hBrush = iif( IsHot, ghExplorer.hBackBrushHot, ghExplorer.hBackBrushNode)
         backclr = iif( IsHot, ghExplorer.BackColorHot, ghExplorer.BackColorNode)
         foreclr = iif( IsHot, ghExplorer.ForeColorNode, ghExplorer.ForeColorNode)
      end if
         
      ' Paint the entire background
      ' Create our rect that works with the entire line
      SetRect( @rc, 0, 0, nWidth, nHeight )
      FillRect( memDC, @rc, hBrush )

      SetBkColor( memDC, backclr )   
      SetTextColor( memDC, foreclr )

      dim as RECT rcText = rc
      dim as RECT rcBitmap = rc

      dim as long wsStyle 
      
      ' indent the text based on its type
      if isNodeHeader then
         rcBitmap.right = rcBitmap.left + pWindow->ScaleX(20)
         SelectObject( memDC, ghMenuBar.hFontSymbol )
         wsStyle = DT_NOPREFIX or DT_CENTER or DT_TOP or DT_SINGLELINE  
         if isIconDown then 
            DrawText( memDC, wszChevronDown, -1, cast(lpRect, @rcBitmap), wsStyle )
         else   
            DrawText( memDC, wszChevronRight, -1, cast(lpRect, @rcBitmap), wsStyle )
         end if
         wszCaption = mid(wszCaption, 6)
         rcText.left = rcBitmap.right
         SelectObject( memDC, ghMenuBar.hFontMenuBar )
         wsStyle = DT_NOPREFIX or DT_LEFT or DT_VCENTER or DT_SINGLELINE  
         DrawText( memDC, wszCaption.sptr, -1, cast(lpRect, @rcText), wsStyle )
      else
         ' This would be a regular file.
         rcBitmap.left = rcText.left + pWindow->ScaleX(20)
         rcBitmap.right = rcBitmap.left + pWindow->ScaleX(20) 

         SelectObject( memDC, ghMenuBar.hFontSymbol )
         wsStyle = DT_NOPREFIX or DT_CENTER or DT_TOP or DT_SINGLELINE  
         DrawText( memDC, wszDocumentIcon, -1, cast(lpRect, @rcBitmap), wsStyle )

         rcText.left = rcBitmap.right
         SelectObject( memDC, ghMenuBar.hFontMenuBar )
         wsStyle = DT_NOPREFIX or DT_LEFT or DT_VCENTER or DT_SINGLELINE  
         DrawText( memDC, wszCaption.sptr, -1, cast(lpRect, @rcText), wsStyle )
      end if
         
      BitBlt lpdis->hDC, lpdis->rcItem.left, lpdis->rcItem.top, _
             nWidth, nHeight, memDC, 0, 0, SRCCOPY 

      ' Cleanup
      if hbit  then DeleteObject SelectObject(memDC, hbit)
      if memDC then DeleteDC memDC
      RestoreDC(lpdis->hDC, -1)
   end if   

   function = true
   
end function


' ========================================================================================
' Process WM_LBUTTONUP message for window/dialog: frmExplorer
' ========================================================================================
function frmExplorer_OnLButtonUp( _
            byval HWnd as HWnd, _
            byval x as long, _
            byval y as long, _
            byval keyflags as UINT _
            ) as long
  
   if gApp.bDragActive then
      gApp.bDragActive = false 
      gApp.hWndPanel = 0
      ReleaseCapture()
   end if
   
   function = 0
end function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmExplorer
' ========================================================================================
function frmExplorer_OnCommand( _
            byval HWnd as HWnd, _
            byval id as long, _
            byval hwndCtl as HWnd, _
            byval codeNotify as UINT _
            ) as LRESULT

   select case codeNotify
      case LBN_SELCHANGE
         ' update the highlighting of the current line
         AfxRedrawWindow(hwndCtl)
         ' update the scrollbar position if necessary
         frmExplorer_PositionWindows()
   end select      

   function = 0
end function


' ========================================================================================
' frmExplorer Window procedure
' ========================================================================================
function frmExplorer_WndProc( _
            byval HWnd   as HWnd, _
            byval uMsg   as UINT, _
            byval wParam as WPARAM, _
            byval lParam as LPARAM _
            ) as LRESULT

   select case uMsg
      HANDLE_MSG (HWnd, WM_SIZE,        frmExplorer_OnSize)
      HANDLE_MSG (HWnd, WM_PAINT,       frmExplorer_OnPaint)
      HANDLE_MSG (HWnd, WM_COMMAND,     frmExplorer_OnCommand)
      HANDLE_MSG (HWnd, WM_LBUTTONDOWN, frmExplorer_OnLButtonDown)
      HANDLE_MSG (HWnd, WM_LBUTTONUP,   frmExplorer_OnLButtonUp)
      HANDLE_MSG (HWnd, WM_MOUSEMOVE,   frmExplorer_OnMouseMove)
      HANDLE_MSG (HWnd, WM_MEASUREITEM, frmExplorer_OnMeasureItem)
      HANDLE_MSG (HWnd, WM_DRAWITEM,    frmExplorer_OnDrawItem)
   
   case WM_ERASEBKGND
      return true

   case WM_CTLCOLORSTATIC
      ' Foreground and Background colors for the Explorer label
      SetTextColor cast(HDC, wParam), ghExplorer.forecolor
      SetBkColor cast(HDC, wParam), ghExplorer.backcolor
      return cast( LRESULT, ghExplorer.hPanelBrush)

   end select

   ' for messages that we don't deal with
   function = DefWindowProc( HWnd, uMsg, wParam, lParam )

end function


' ========================================================================================
' TrackMouse. setup TrackMouseEvent for a specififc window.
' ========================================================================================
function frmExplorer_TrackMouse( byval hWin as HWND ) as long
   dim tme as TrackMouseEvent
   tme.cbSize = sizeof(TrackMouseEvent)
   tme.dwFlags = TME_HOVER or TME_LEAVE
   tme.hwndTrack = hWin
   TrackMouseEvent(@tme) 
   function = 0
end function         

' ========================================================================================
' frmExplorerListBox_SubclassProc 
' ========================================================================================
function frmExplorerListBox_SubclassProc ( _
            byval hWin   as HWnd, _                 ' // Control window handle
            byval uMsg   as UINT, _                 ' // Type of message
            byval _wParam as WPARAM, _               ' // First message parameter
            byval _lParam as LPARAM, _               ' // Second message parameter
            byval uIdSubclass as UINT_PTR, _        ' // The subclass ID
            byval dwRefData as DWORD_PTR _          ' // Pointer to reference data
            ) as LRESULT

   dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMEXPLORER)
   static as long accumDelta
   
   ' keep track of last index we were over so that we only issue a 
   ' repaint if the cursor has moved off of the line
   static as long nLastIdx = -1
   static as boolean isTracking = false
      
   select case uMsg
      case MSG_USER_LOAD_EXPLORERFILES
         LoadExplorerFiles()

      Case WM_MOUSEWHEEL
         ' accumulate delta until scroll one line (up +120, down -120). 
         ' 120 is the Microsoft default delta
         dim as long zDelta = GET_WHEEL_DELTA_WPARAM( _wParam )
         dim as long nTopIndex = SendMessage( hWin, LB_GETTOPINDEX, 0, 0 ) 
         accumDelta = accumDelta + zDelta
         if accumDelta >= 120 then       ' scroll up 3 lines
            nTopIndex = nTopIndex - 3
            SendMessage( hWin, LB_SETTOPINDEX, nTopIndex, 0 ) 
            accumDelta = 0
            frmExplorerVScroll_PositionWindows( SW_SHOWNA )
         elseif accumDelta <= -120 then  ' scroll down 3 lines
            nTopIndex = nTopIndex + 3
            SendMessage( hWin, LB_SETTOPINDEX, nTopIndex, 0 ) 
            accumDelta = 0
            frmExplorerVScroll_PositionWindows( SW_SHOWNA )
         end if
         
      Case WM_MOUSEMOVE
         ' Track that we are over the control in order to catch the 
         ' eventual WM_MOUSELEAVE event
         if not isTracking then
            frmExplorer_TrackMouse( hWin )
            isTracking = true
         end if
         
         ' get the item rect that the mouse is over and only invalidate
         ' that instead of the entire listbox
         dim as RECT rc
         dim as long idx = Listbox_ItemFromPoint( hWin, GET_X_LPARAM(_lParam), GET_Y_LPARAM(_lParam)) 
         if idx <> nLastIdx then
            ListBox_GetItemRect( hWin, idx, @rc )
            InvalidateRect( hWin, @rc, true )
            ListBox_GetItemRect( hWin, nLastIdx, @rc )
            InvalidateRect( hWin, @rc, true )
            nLastIdx = idx
         end if
                     
      case WM_MOUSELEAVE
         isTracking = false   
         nLastIdx = -1
         AfxRedrawWindow(hWin)   
        
         
      case WM_RBUTTONDOWN
         ' Create the popup menu
         dim as long idx = Listbox_ItemFromPoint( hWin, GET_X_LPARAM(_lParam), GET_Y_LPARAM(_lParam)) 
         dim as clsDocument ptr pDoc = cast(clsDocument ptr, ListBox_GetItemData( hWin, idx ))
         if pDoc then
            ListBox_SetSel( hWin, true, idx )
            dim as HMENU hPopupMenu = CreateExplorerContextMenu(pDoc)
            dim as POINT pt: GetCursorPos( @pt )
            dim as long id = TrackPopupMenu(hPopUpMenu, TPM_RETURNCMD, pt.x, pt.y, 0, HWND_FRMMAIN, byval null)
            ' process the selected rows. Need to save the current state of selected
            ' rows prior to processing because as the items are processed some of the
            ' actions may reset the state of listbox. We save the pDoc's in an array
            ' rather than index values because the listbox may get reloaded during
            ' some actions invalidating some indexes. 
            dim as long nSelCount = SendMessage( hWin, LB_GETSELCOUNT, 0, 0 )
            dim pDocArray(nSelCount-1) as clsDocument ptr 
            dim as long nextIdx = 0
            for i as long = 0 to ListBox_GetCount( hWin ) - 1
               if ListBox_GetSel( hWin, i ) then
                  dim as CWSTR wszCaption = AfxGetListBoxText( hWin, i )
                  if left(wszCaption, 1) = "%" then 
                     pDocArray(nextIdx) = 0
                  else   
                     pDocArray(nextIdx) = cast(clsDocument ptr, ListBox_GetItemData( hWin, i ))
                  end if
                  nextIdx = nextIdx + 1
               end if
            next         
            ' Process our array of selected pDoc's.
            for i as long = lbound(pDocArray) to ubound(pDocArray)
               if pDocArray(i) = 0 then continue for
               select case id
                  case IDM_FILEOPEN_EXPLORERLISTBOX, IDM_FILECLOSE_EXPLORERLISTBOX, _
                       IDM_FILESAVE_EXPLORERLISTBOX, IDM_FILESAVEAS_EXPLORERLISTBOX
                       OnCommand_FileExplorerMessage( id, pDocArray(i) )
                  case IDM_REMOVEFILEFROMPROJECT_EXPLORERLISTBOX
                       OnCommand_ProjectRemove( id, pDocArray(i) )
                  case IDM_SETFILEMAIN_EXPLORERTREEVIEW, IDM_SETFILERESOURCE_EXPLORERTREEVIEW, _
                       IDM_SETFILEHEADER_EXPLORERTREEVIEW, IDM_SETFILEMODULE_EXPLORERTREEVIEW, _
                       IDM_SETFILENORMAL_EXPLORERTREEVIEW 
                       OnCommand_ProjectSetFileType( id, pDocArray(i) )
               end select
            next
            LoadExplorerFiles()
            frmExplorer_UnSelectListBox() 
            DestroyMenu( hPopUpMenu )
            Return true   ' prevent further processing that leads to WM_CONTEXTMENU
         end if
         
      case WM_LBUTTONUP
         ' Prevent this programmatic selection if Ctrl or Shift is active
         ' because we want the listbox to select the listbox item rather for
         ' us to mess with that selection via SetTabIndexByDocumentPtr().
         dim as boolean isCtrl = (GetAsyncKeyState(VK_CONTROL) and &H8000)
         dim as boolean isShift = (GetAsyncKeyState(VK_SHIFT) and &H8000) 
         if (isCtrl = false) andalso (isShift = false) then
            ' determine if we clicked on a regular file or a node header
            dim as RECT rc
            dim as long idx = Listbox_ItemFromPoint( hWin, GET_X_LPARAM(_lParam), GET_Y_LPARAM(_lParam)) 
            dim as CWSTR wszCaption = AfxGetListBoxText( hWin, idx )
            if left(wszCaption, 1) = "%" then
               ' Toggle the show/hide of files under this node
               dim as long idxArray = val(mid(wszCaption, 2, 1))
               gExplorerNodeShow(idxArray) = not gExplorerNodeShow(idxArray)
               ' allow listbox click event to fully process before loading new explorer files
               ' so that we can correctly select the current item.
               PostMessage( hWin, MSG_USER_LOAD_EXPLORERFILES, 0, 0 )
            else
               ' if the file is already showing in the Top tabs then switch to that tab.
               ' We do not open a new tab on single click. Only double click or ENTER will
               ' open a new tab. 
               dim as clsDocument ptr pDoc = cast(clsDocument ptr, ListBox_GetItemData( hWin, idx ))
               gTTabCtl.SetTabIndexByDocumentPtr( pDoc )
            end if
         end if
                  
      case WM_LBUTTONDBLCLK 
         ' determine if we clicked on a regular file or a node header. if it is a regular
         ' file then load it.
         dim as RECT rc
         dim as long idx = Listbox_ItemFromPoint( hWin, GET_X_LPARAM(_lParam), GET_Y_LPARAM(_lParam)) 
         dim as CWSTR wszCaption = AfxGetListBoxText( hWin, idx )
         if left(wszCaption, 1) <> "%" then
            dim as CWSTR wszFilename
            dim as clsDocument ptr pDoc = cast(clsDocument ptr, ListBox_GetItemData( hWin, idx ))
            if pDoc then
               frmFunctionList_GetFileNameFunctionName( pDoc->hNodeExplorer, wszFilename, "" )
               OpenSelectedDocument( wszFilename, "" )
            end if   
         end if
         
      case WM_ERASEBKGND
         ' if the number of lines in the listbox maybe less than the number per page then 
         ' calculate from last item to bottom of listbox, otherwise calculate based on
         ' the mod of the lineheight to listbox height so we can color the partial line
         ' that won't be displayed at the bottom of the list.
         dim as RECT rc: GetClientRect( hWin, @rc )

         dim as RECT rcItem  
         SendMessage( hWin, LB_GETITEMRECT, 0, cast(LPARAM, @rcItem) )
         dim as long itemHeight = rcItem.bottom - rcItem.top
         dim as long NumItems = ListBox_GetCount(hWin)
         dim as long nTopIndex = SendMessage( hWin, LB_GETTOPINDEX, 0, 0 ) 
         dim as long visible_rows = 0
         dim as long ItemsPerPage = 0
         dim as long bottom_index = 0
                     
         if NumItems > 0 then
            ItemsPerPage = (rc.bottom - rc.top) / itemHeight
            bottom_index = (nTopIndex + ItemsPerPage) 
            if bottom_index >= NumItems then bottom_index = NumItems - 1
            visible_rows = (bottom_index - nTopIndex) + 1
            rc.top = visible_rows * itemHeight 
         end if

         if rc.top < rc.bottom then
            dim as HDC _hDC = cast(HDC, _wParam)
            FillRect( _hDC, @rc, ghExplorer.hPanelBrush )
         end if

         ValidateRect( hWin, @rc )
         return true
      
      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass( hWin, @frmExplorerListBox_SubclassProc, uIdSubclass )
   end select
    
   ' For messages that we don't deal with
   function = DefSubclassProc(hWin, uMsg, _wParam, _lParam)

end function


' ========================================================================================
' frmExplorer_Show
' ========================================================================================
function frmExplorer_Show( byval hWndParent as HWnd ) as LRESULT

   '  Create the main window and child controls
   dim pWindow as CWindow ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   ' Only make the Explorer initially visible if it was already visible
   ' when the most previous instance of the program closed. Also, set the width of
   ' the window to the last used visible width.
   dim as long nWidth = iif(gConfig.ShowExplorer, gConfig.ShowExplorerWidth, 250)
   
   HWND_FRMEXPLORER = pWindow->Create( hWndParent, "Explorer Window", @frmExplorer_WndProc, 0, 0, nWidth, 0, _
        WS_CHILD or iif(gConfig.ShowExplorer, WS_VISIBLE, 0) or WS_CLIPSIBLINGS or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT or WS_EX_LEFT or WS_EX_LTRREADING or WS_EX_RIGHTSCROLLBAR)

   ' Disable background erasing by only assigning the one style
   pWindow->ClassStyle = CS_DBLCLKS
   pWindow->SetFont( "Segoe UI Semibold", 9)

   dim as HWnd hBtn = _
   pWindow->AddControl("LABEL", , IDC_FRMEXPLORER_LBLEXPLORER, L(247,"EXPLORER"), 10, 12, 75, 20, _
        WS_CHILD or WS_VISIBLE or WS_CLIPSIBLINGS or WS_CLIPCHILDREN or SS_LEFT, _
        WS_EX_LEFT or WS_EX_LTRREADING)
   pWindow->Brush = ghExplorer.hPanelBrush

   HWND_FRMEXPLORER_LISTBOX = _ 
        pWindow->AddControl("LISTBOX", , IDC_FRMEXPLORER_LISTBOX, "", 0, 0, 0, 0, _
        WS_CHILD or WS_CLIPSIBLINGS or WS_CLIPCHILDREN or WS_TABSTOP or _
        LBS_NOINTEGRALHEIGHT or LBS_EXTENDEDSEL or LBS_MULTIPLESEL or _
        LBS_OWNERDRAWFIXED or LBS_HASSTRINGS or LBS_NOTIFY, _
        WS_EX_LEFT or WS_EX_RIGHTSCROLLBAR, , _
        cast(SUBCLASSPROC, @frmExplorerListBox_SubclassProc), _
        IDC_FRMEXPLORER_LISTBOX, cast(DWORD_PTR, @pWindow))

   function = 0
   
end function


' ========================================================================================
' LoadExplorerFiles
' This will clear the current list of files in the Explorer and repopulate it
' with the latest list of filenames that are stored in the hidden FunctionList treeview.
' Files are excluded from being added based on the state of the global array gExplorerNodeShow()
' ========================================================================================
function LoadExplorerFiles() as long
   dim as HWND hList = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_LISTBOX)
   dim as Hwnd hTree = GetDlgItem(HWND_FRMFUNCTIONLIST, IDC_FRMFUNCTIONLIST_TREE)
    
   ' Hide the listbox while it is loading so that we don't get the unpainted
   ' white background from the empty listbox
   ShowWindow( hList, SW_HIDE )
   
   ' Save the topindex because we will restore it after filling the new contents 
   dim as long nTopIndex = SendMessage( hList, LB_GETTOPINDEX, 0, 0 ) 

   dim pDocActive as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   
   dim as long idxRestore = -1

   ' Clear all content from the listbox
   ListBox_ResetContent(hList)
   
   ' Add the list of filenames (allow the treeview to sort the data for us)
   dim as HTREEITEM hParent
   dim as long p
   dim wszText as wstring * MAX_PATH

   redim hRootNodes(any) as HTREEITEM
   
   '#define NODETYPE_MAIN              0
   '#define NODETYPE_RESOURCE          1
   '#define NODETYPE_HEADER            2
   '#define NODETYPE_MODULE            3
   '#define NODETYPE_NORMAL            4

   if gApp.IsProjectActive then
      redim hRootNodes(4) as HTREEITEM
      hRootNodes(0) = gApp.hExplorerMainNode          
      hRootNodes(1) = gApp.hExplorerResourceNode      
      hRootNodes(2) = gApp.hExplorerHeaderNode        
      hRootNodes(3) = gApp.hExplorerModuleNode        
      hRootNodes(4) = gApp.hExplorerNormalNode        
   else   
      redim hRootNodes(0) as HTREEITEM
      hRootNodes(0) = gApp.hExplorerFilesNode
   end if
   
   for ii as long = lbound(hRootNodes) to ubound(hRootNodes)

      ' The node header description starts with a % in order for the
      ' drawing routine to know that it is a node header and not a
      ' regular file.
      ' e.g.
      '       %0%1%Main
      '       %1%1%Resource
      '       %2%0%Header
      '       %3%0%Module
      '       %4%1%Normal
      ' The value after the % is the array index into gExplorerNodeShow
      ' that allows the mouse routines to toggle on/off displaying of
      ' the files. The value after the second % indicates to the drawing 
      ' routine that files exist ("1") under the node, or ("0") that 
      ' no files exist. This wil then dictate what icon to show next
      ' to the line.
      
      wszText = "%" & ii & "%"
      if gExplorerNodeShow(ii) = false then 
         wszText = wszText & "0" & "%"
      else   
         wszText = wszText & "1" & "%"
      end if

      select case ii
         case NODETYPE_MAIN
            if gApp.IsProjectActive then
               wszText = wszText & L(212, "Main")
            else   
               wszText = wszText & L(2, "File")
            end if   
         case NODETYPE_RESOURCE: wszText = wszText & L(213, "Resource")
         case NODETYPE_HEADER:   wszText = wszText & L(175, "Header")
         case NODETYPE_MODULE:   wszText = wszText & L(211, "Module")
         case NODETYPE_NORMAL:   wszText = wszText & L(210, "Normal")
      end select
      
      p = Listbox_AddString( hList, @wszText )

      ' Check the global array gExplorerNodeShow to see if these types
      ' of files should be loaded into the listbox.
      if gExplorerNodeShow(ii) = false then continue for
      
      Treeview_SortChildren( hTree, hRootNodes(ii), 0 )
      hParent = Treeview_GetChild( hTree, hRootNodes(ii))
      do until hParent = 0 
         ' Filename
         Treeview_GetItemText(hTree, hParent, @wszText, MAX_PATH)
         dim as clsDocument ptr pDoc = cast(clsDocument ptr, TreeView_GetlParam(hTree, hParent))
         p = Listbox_AddString( hList, @wszText )
         ' pDoc item is stored in ListBox ItemData
         ListBox_SetItemData( hList, p, pDoc )    
         if pDoc then pDoc->hNodeExplorer = hParent
         if pDocActive then 
            if hParent = pDocActive->hNodeExplorer then idxRestore = p
         end if
         ' Get the next filename
         hParent = Treeview_GetNextSibling(hTree, hParent)
      loop 
   
   next

   ' Restore the top index so the list displays like it did before being reset
   SendMessage( hList, LB_SETTOPINDEX, nTopIndex, 0 ) 

   ' Ensure that Listbox is now properly sized and then show 
   ' the listbox now that it is fully populated
   ShowWindow( hList, SW_SHOW )
   frmExplorer_PositionWindows()

   ' Set current selection to the item that was selected  prior
   ' to us reloading the listbox
   ListBox_SetSel( hList, true, idxRestore )
   AfxRedrawWindow(hList)

   ' Determine if the VScroll bar has changed size or is now hidden/shown
   frmExplorerVScroll_PositionWindows( SW_HIDE )

   function = 0
end function

   
