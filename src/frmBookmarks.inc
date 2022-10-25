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
''  frmBookmarks.inc
''   

#include once "frmBookmarks.bi"

' ========================================================================================
' Get the Bookmarks line number from the Listbox line
' ========================================================================================
function getBookmarksLineNumber( byval wszCaption as CWSTR ) as long
   ' do not use Parse for this because line may contain embedded % in description
   dim as long nLineNum 
   dim as long f1
   f1 = instr(wszCaption, "%")
   if f1 then nLineNum = val(left(wszCaption, f1-1))
   function = nLineNum
end function

' ========================================================================================
' Get the Bookmarks description from the Listbox line
' ========================================================================================
function getBookmarksDescription( byval wszCaption as CWSTR ) as CWSTR
   ' do not use Parse for this because line may contain embedded % in description
   dim as CWSTR wszTemp = "Error retrieving bookmark"
   dim as long f1
   f1 = instr(wszCaption, "%")
   if f1 then wszTemp = mid(wszCaption, f1 + 1)
   function = wszTemp
end function


' ========================================================================================
' Expand/Collapse all Bookmark Nodes
' ========================================================================================
function frmBookmarks_ExpandAll() as long
   dim pDoc as clsDocument ptr = gApp.pDocList
   do until pDoc = 0
      pDoc->bBookmarkExpanded = true
      pDoc = pDoc->pDocNext
   loop
   LoadBookmarksFiles()
   function = 0
end function

function frmBookmarks_CollapseAll() as long
   dim pDoc as clsDocument ptr = gApp.pDocList
   do until pDoc = 0
      pDoc->bBookmarkExpanded = false
      pDoc = pDoc->pDocNext
   loop
   LoadBookmarksFiles()
   function = 0
end function   


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
function frmBookmarks_PositionWindows() as LRESULT

   ' Get the entire client area
   dim as Rect rc
   GetClientRect( HWND_FRMBOOKMARKS, @rc )

   SetWindowPos( HWND_FRMBOOKMARKS_LISTBOX, 0, _
                 rc.left, rc.top, rc.right-rc.left, rc.bottom-rc.top, _
                 SWP_NOZORDER )

   function = 0
end function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmBookmarks
' ========================================================================================
function frmBookmarks_OnSize( _
            byval HWnd as HWnd, _
            byval state as UINT, _
            byval cx as long, _
            byval cy as long _
            ) as LRESULT
   if state <> SIZE_MINIMIZED then
      ' Position all of the child windows
      frmBookmarks_PositionWindows
   end if
   function = 0
end function
       
' ========================================================================================
' Process WM_PAINT message for window/dialog: frmBookmarks
' ========================================================================================
function frmBookmarks_OnPaint( byval HWnd as HWnd ) as LRESULT
            
   dim as PAINTSTRUCT ps
   dim as HDC hDc
   
   hDC = BeginPaint(hWnd, @ps)

   SaveDC( hDC )
   FillRect( hDC, @ps.rcPaint, ghPanel.hPanelBrush )
   RestoreDC( hDC, -1 )
   EndPaint( hWnd, @ps )
   
   function = 0
end function


' ========================================================================================
' Process WM_MEASUREITEM message for window/dialog: frmBookmarks
' ========================================================================================
function frmBookmarks_OnMeasureItem( _
            byval HWnd as HWnd, _
            byval lpmis as MEASUREITEMSTRUCT ptr _
            ) as long
   ' Set the height of the list box items. 
   dim pWindow as CWindow ptr = AfxCWindowPtr(HWnd)
   lpmis->itemHeight = pWindow->ScaleY(EXPLORERITEM_HEIGHT)
      
   function = 0
end function


' ========================================================================================
' Process WM_DRAWITEM message for window/dialog: frmBookmarks
' ========================================================================================
function frmBookmarks_OnDrawItem( _
            byval HWnd as HWnd, _
            byval lpdis as const DRAWITEMSTRUCT ptr _
            ) as long

   dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMMAIN)
   if pWindow = 0 then exit function

   if lpdis = 0 then exit function
 
   if ( lpdis->itemAction = ODA_DRAWENTIRE ) orelse _
      ( lpdis->itemAction = ODA_SELECT ) orelse _
      ( lpdis->itemAction = ODA_FOCUS ) then

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
      dim as HBRUSH hBrush = ghPanel.hBackBrush
      dim as COLORREF foreclr = ghPanel.ForeColor
      dim as COLORREF backclr = ghPanel.BackColor
         
      dim as boolean IsHot = false
      dim as boolean isNodeHeader = false
      dim as boolean isIconDown = false
         
      dim as POINT pt
      GetCursorPos( @pt )
      MapWindowPoints( lpdis->hwndItem, HWND_DESKTOP, cast( POINT ptr, @rc ), 2 )
      if PtInRect( @rc, pt ) then IsHot = true
      
      ' if mouse is over VScrollBar then reset hot
      if isMouseOverWindow( HWND_FRMPANEL_VSCROLLBAR ) then IsHot = false
         
      if ListBox_GetCurSel(lpdis->hwndItem) = lpdis->itemID then IsHot = true

      hBrush = iif( IsHot, ghPanel.hBackBrushHot, ghPanel.hBackBrush)
      backclr = iif( IsHot, ghPanel.BackColorHot, ghPanel.BackColor)
      foreclr = iif( IsHot, ghPanel.ForeColorHot, ghPanel.ForeColor)

      dim as CWSTR wszCaption = AfxGetListBoxText(lpdis->hwndItem, lpdis->ItemID)
      dim as clsDocument ptr pDoc = cast(clsDocument ptr, lpdis->itemData)

      ' if this is a "node" header 
      if left(wszCaption, 4) = "true" then 
         isNodeHeader = true
         isIconDown = true
         if pDoc then wszCaption = AfxStrPathName( "NAMEX", pDoc->DiskFilename )
      elseif left(wszCaption, 5) = "false" then 
         isNodeHeader = true
         isIconDown = false
         if pDoc then wszCaption = AfxStrPathName( "NAMEX", pDoc->DiskFilename )
      else
         ' must be a bookmark line 
          wszCaption = getBookmarksDescription( wszCaption )
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
         wszCaption = wszCaption
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
         
      BitBlt( lpdis->hDC, lpdis->rcItem.left, lpdis->rcItem.top, _
                nWidth, nHeight, memDC, 0, 0, SRCCOPY )

      ' Cleanup
      if hbit  then DeleteObject SelectObject(memDC, hbit)
      if memDC then DeleteDC memDC
      RestoreDC(lpdis->hDC, -1)
   end if   

   function = true
   
end function

' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmBookmarks
' ========================================================================================
function frmBookmarks_OnCommand( _
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
         frmBookmarks_PositionWindows()
   end select      

   function = 0
end function


' ========================================================================================
' frmBookmarks Window procedure
' ========================================================================================
function frmBookmarks_WndProc( _
            byval HWnd   as HWnd, _
            byval uMsg   as UINT, _
            byval wParam as WPARAM, _
            byval lParam as LPARAM _
            ) as LRESULT

   static hTooltip as HWND

   select case uMsg
      HANDLE_MSG (HWnd, WM_SIZE,        frmBookmarks_OnSize)
      HANDLE_MSG (HWnd, WM_PAINT,       frmBookmarks_OnPaint)
      HANDLE_MSG (HWnd, WM_COMMAND,     frmBookmarks_OnCommand)
      HANDLE_MSG (HWnd, WM_MEASUREITEM, frmBookmarks_OnMeasureItem)
      HANDLE_MSG (HWnd, WM_DRAWITEM,    frmBookmarks_OnDrawItem)
   
   case WM_ERASEBKGND
      return true

   end select
   
   ' for messages that we don't deal with
   function = DefWindowProc( HWnd, uMsg, wParam, lParam )

end function

' ========================================================================================
' frmBookmarksListBox_SubclassProc 
' ========================================================================================
function frmBookmarksListBox_SubclassProc ( _
            byval hWin   as HWnd, _                 ' // Control window handle
            byval uMsg   as UINT, _                 ' // Type of message
            byval _wParam as WPARAM, _               ' // First message parameter
            byval _lParam as LPARAM, _               ' // Second message parameter
            byval uIdSubclass as UINT_PTR, _        ' // The subclass ID
            byval dwRefData as DWORD_PTR _          ' // Pointer to reference data
            ) as LRESULT

   dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMBOOKMARKS)
   static as long accumDelta
   static as HWND hTooltip
      
   ' keep track of last index we were over so that we only issue a 
   ' repaint if the cursor has moved off of the line
   static as long nLastIdx = -1
      
   select case uMsg
      case MSG_USER_LOAD_BOOKMARKSFILES
         LoadBookmarksFiles()

      Case WM_MOUSEWHEEL
         ' accumulate delta until scroll one line (up +120, down -120). 
         ' 120 is the Microsoft default delta
         dim as long zDelta = GET_WHEEL_DELTA_WPARAM( _wParam )
         dim as long nTopIndex = SendMessage( hWin, LB_GETTOPINDEX, 0, 0 ) 
         accumDelta = accumDelta + zDelta
         if accumDelta >= 120 then       ' scroll up 3 lines
            nTopIndex = nTopIndex - 3
            nTopIndex = max( 0, nTopIndex )
            SendMessage( hWin, LB_SETTOPINDEX, nTopIndex, 0 ) 
            accumDelta = 0
            frmPanelVScroll_PositionWindows( SW_SHOWNA )
         elseif accumDelta <= -120 then  ' scroll down 3 lines
            nTopIndex = nTopIndex + 3
            SendMessage( hWin, LB_SETTOPINDEX, nTopIndex, 0 ) 
            accumDelta = 0
            frmPanelVScroll_PositionWindows( SW_SHOWNA )
         end if
         
      Case WM_MOUSEMOVE
         ' Track that we are over the control in order to catch the 
         ' eventual WM_MOUSEHOVER and WM_MOUSELEAVE events
         dim tme as TrackMouseEvent
         tme.cbSize = sizeof(TrackMouseEvent)
         tme.dwFlags = TME_HOVER or TME_LEAVE
         tme.hwndTrack = hWin
         TrackMouseEvent(@tme) 
         
         ' get the item rect that the mouse is over and only invalidate
         ' that instead of the entire listbox
         dim as RECT rc
         dim as long idx = Listbox_ItemFromPoint( hWin, GET_X_LPARAM(_lParam), GET_Y_LPARAM(_lParam)) 
         ' The return value contains the index of the nearest item in the LOWORD. The HIWORD is zero 
         ' if the specified point is in the client area of the list box, or one if it is outside the 
         ' client area.
         if hiword(idx) <> 1 then
            if idx <> nLastIdx then
               ListBox_GetItemRect( hWin, idx, @rc )
               InvalidateRect( hWin, @rc, true )
               ListBox_GetItemRect( hWin, nLastIdx, @rc )
               InvalidateRect( hWin, @rc, true )
               nLastIdx = idx
            end if
        end if
                           
      case WM_MOUSEHOVER
         dim as CWSTR wszTooltip 
         if IsWindow(hTooltip) = 0 then hTooltip = AfxAddTooltip( hWin, "", false, false )
         dim as long idx = Listbox_ItemFromPoint( hWin, GET_X_LPARAM(_lParam), GET_Y_LPARAM(_lParam)) 
         ' The return value contains the index of the nearest item in the LOWORD. The HIWORD is zero 
         ' if the specified point is in the client area of the list box, or one if it is outside the 
         ' client area.
         if hiword(idx) <> 1 then
            dim as CWSTR wszCaption = AfxGetListBoxText( hWin, idx )
            if (left(wszCaption, 4) = "true") orelse (left(wszCaption, 5) = "false") then
               dim as clsDocument ptr pDoc = cast(clsDocument ptr, ListBox_GetItemData( hWin, idx ))
               if pDoc then wszTooltip = pDoc->DiskFilename
               ' Display the tooltip
               AfxSetTooltipText( hTooltip, hWin, wszTooltip )
               AfxRedrawWindow( hWin )
            end if
         end if
         
      case WM_MOUSELEAVE
         nLastIdx = -1
         AfxRedrawWindow(hWin)   
         
      case WM_RBUTTONDOWN
         ' Create the popup menu
         dim as long idx = Listbox_ItemFromPoint( hWin, GET_X_LPARAM(_lParam), GET_Y_LPARAM(_lParam)) 
         ' The return value contains the index of the nearest item in the LOWORD. The HIWORD is zero 
         ' if the specified point is in the client area of the list box, or one if it is outside the 
         ' client area.
         if hiword(idx) <> 1 then
            dim as clsDocument ptr pDoc = cast(clsDocument ptr, ListBox_GetItemData( hWin, idx ))
            if pDoc then
               dim as HMENU hPopupMenu
               dim as CWSTR wszCaption = AfxGetListBoxText( hWin, idx )
               if (left(wszCaption, 4) = "true") orelse (left(wszCaption, 5) = "false") then
                  hPopupMenu = CreateBookmarksHeaderNodeContextMenu()
               else
                  hPopupMenu = CreateBookmarksBookmarkNodeContextMenu()   
               end if
               dim as POINT pt: GetCursorPos( @pt )
               dim as long id = TrackPopupMenu(hPopUpMenu, TPM_RETURNCMD, pt.x, pt.y, 0, HWND_FRMMAIN, byval null)
               select case id
                  case IDM_CLEARALLBOOKMARKNODE
                     SciExec( pDoc->hWindow(0), SCI_MARKERDELETEALL, -1, 0 )
                  case IDM_REMOVEBOOKMARKNODE
                     dim as long nLineNum = val(AfxStrParse(wszCaption, 1, "%"))
                     pDoc->ToggleBookmark( nLineNum )
               end select
               LoadBookmarksFiles()
               DestroyMenu( hPopUpMenu )
               Return true   ' prevent further processing that leads to WM_CONTEXTMENU
            end if
         end if
         
      case WM_LBUTTONUP
         ' determine if we clicked on a regular file or a node header
         dim as RECT rc
         dim as long idx = Listbox_ItemFromPoint( hWin, GET_X_LPARAM(_lParam), GET_Y_LPARAM(_lParam)) 
         ' The return value contains the index of the nearest item in the LOWORD. The HIWORD is zero 
         ' if the specified point is in the client area of the list box, or one if it is outside the 
         ' client area.
         if hiword(idx) <> 1 then
            dim as clsDocument ptr pDoc = cast(clsDocument ptr, ListBox_GetItemData( hWin, idx ))
            dim as CWSTR wszCaption = AfxGetListBoxText( hWin, idx )
            if (left(wszCaption, 4) = "true") orelse (left(wszCaption, 5) = "false") then
               ' Toggle the show/hide of bookmarks under this node
               if pDoc then pDoc->bBookmarkExpanded = not pDoc->bBookmarkExpanded
               ' allow listbox click event to fully process before loading new bookmarks
               PostMessage( hWin, MSG_USER_LOAD_BOOKMARKSFILES, 0, 0 )
            else
               ' Attempt to show the bookmark
               dim as long nLineNum = getBookmarksLinenumber( wszCaption )
               dim as CWSTR wszDiskFilename 
               if pDoc then wszDiskFilename = pDoc->DiskFilename
               OpenSelectedDocument( wszDiskFilename, "", nLineNum )
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
            FillRect( _hDC, @rc, ghPanel.hPanelBrush )
         end if

         ValidateRect( hWin, @rc )
         return true
      
      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass( hWin, @frmBookmarksListBox_SubclassProc, uIdSubclass )
   end select
    
   ' For messages that we don't deal with
   function = DefSubclassProc( hWin, uMsg, _wParam, _lParam )

end function


' ========================================================================================
' frmBookmarks_Show
' ========================================================================================
function frmBookmarks_Show( byval hWndParent as HWnd ) as LRESULT

   '  Create the main window and child controls
   dim pWindow as CWindow ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   HWND_FRMBOOKMARKS = pWindow->Create( hWndParent, "Bookmarks Window", @frmBookmarks_WndProc, _
        0, 0, 0, 0, _
        WS_CHILD or WS_CLIPSIBLINGS or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT or WS_EX_LEFT or WS_EX_LTRREADING or WS_EX_RIGHTSCROLLBAR)

   ' Disable background erasing by only assigning the one style
   pWindow->ClassStyle = CS_DBLCLKS

   HWND_FRMBOOKMARKS_LISTBOX = _ 
        pWindow->AddControl("LISTBOX", , IDC_FRMBOOKMARKS_LISTBOX, "", 0, 0, 0, 0, _
        WS_CHILD or WS_CLIPSIBLINGS or WS_CLIPCHILDREN or WS_TABSTOP or _
        LBS_NOINTEGRALHEIGHT or LBS_OWNERDRAWFIXED or LBS_HASSTRINGS or LBS_NOTIFY, _
        WS_EX_LEFT or WS_EX_RIGHTSCROLLBAR, , _
        cast(SUBCLASSPROC, @frmBookmarksListBox_SubclassProc), _
        IDC_FRMBOOKMARKS_LISTBOX, cast(DWORD_PTR, @pWindow))

   function = 0
   
end function


' ========================================================================================
' LoadBookmarksFiles
' This will clear the current list of files in the listbox and repopulate it
' with the latest list of filenames that are stored in the hidden FunctionList treeview.
' ========================================================================================
function LoadBookmarksFiles() as long
   dim as HWND hList = GetDlgItem(HWND_FRMBOOKMARKS, IDC_FRMBOOKMARKS_LISTBOX)
    
   ' Hide the listbox while it is loading so that we don't get the unpainted
   ' white background from the empty listbox
   ShowWindow( hList, SW_HIDE )
   
   ' Save the topindex because we will restore it after filling the new contents 
   dim as long nTopIndex = SendMessage( hList, LB_GETTOPINDEX, 0, 0 ) 

   ' Clear all content from the listbox
   ListBox_ResetContent(hList)
   
   dim wszText as wstring * MAX_PATH
   
   ' Iterate all pDoc in the project/files list
   dim pDoc as clsDocument ptr = gApp.pDocList
   do until pDoc = 0
      ' Filename node line format. The drawing routine will parse this string to determine
      ' what glyph to display as well as the filename only portion of the diskfilename.
      ' true     ' expanded node
      ' false    ' collapsed node
      ' Bookmark node line format. 
      ' 231%Select Case nCount     ' bookmark on line 231 of the specified file

      ' Test if the pDoc has active bookmarks
      dim as long idx, nLineNum, nCount
      dim as string sBookmarks 
      
      ' Search documents that have been loaded and have valid Scintilla window. 
      if pDoc->hWindow(0) <> null then
         sBookmarks = pDoc->GetBookmarks()
      end if
      
      if len(sBookmarks) then
         wszText = wstr(pDoc->bBookmarkExpanded)
         idx = Listbox_AddString( hList, @wszText )
         ListBox_SetItemData( hList, idx, pDoc )

         if pDoc->bBookmarkExpanded then
            nCount = AfxStrParseCount( sBookmarks, "," )
            for i as long = 1 to nCount
               nLineNum = val( AfxStrParse( sBookmarks, i, "," ) )
               wszText = wstr(nLineNum) & _
                         "%" & ltrim(pDoc->GetLine(nLineNum))
               idx = Listbox_AddString( hList, @wszText )
               ListBox_SetItemData( hList, idx, pDoc )
            next
         end if    
      end if
            
      pDoc = pDoc->pDocNext
   loop
   
   ' Restore the top index so the list displays like it did before being reset
   SendMessage( hList, LB_SETTOPINDEX, nTopIndex, 0 ) 

   ' Ensure that Listbox is now properly sized and then show 
   ' the listbox now that it is fully populated (only if it contains any
   ' items because zero items can produce white background).
   if ListBox_GetCount( hList ) then ShowWindow( hList, SW_SHOW )
   frmBookmarks_PositionWindows()

   AfxRedrawWindow( hList )

   ' Determine if the VScroll bar has changed size or is now hidden/shown
   frmPanelVScroll_PositionWindows( SW_HIDE )

   function = 0
end function

   
