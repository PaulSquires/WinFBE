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


' ========================================================================================
' Show or hide the drop down shadow that will display below the top tabs control. The
' width of the shadow is also influenced by whether the Find/Replace dialog is active.
' ========================================================================================
function frmTopTabs_ShowShadow() as long
   
   if IsWindowVisible( HWND_FRMMAIN_TOPTABS ) = 0 then exit function
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   
   if pDoc then
      dim as RECT rcFindReplace = AfxGetWindowRect( HWND_FRMFINDREPLACE )
      dim as RECT rcTopTabs = AfxGetWindowRect( HWND_FRMMAIN_TOPTABS )
      dim as long nLeftWidth = rcTopTabs.right - rcTopTabs.left

      dim as long nLine = SciExec( pDoc->hWindow(0), SCI_GETFIRSTVISIBLELINE, 0, 0)
      dim as long nShow = iif( nLine > 0, SW_SHOWNA, SW_HIDE )
      
      ' if the Find/Replace popup window is active then we always show the show
      ' the shadow window is a popup style so need to use screen coordinates
      if IsWindowVisible( HWND_FRMFINDREPLACE ) then 
         nShow = SW_SHOWNA
         nLeftWidth = rcFindReplace.left - rcTopTabs.left 
      end if
            
      SetWindowPos( HWND_FRMMAIN_TOPTABS_SHADOW, 0, _
                    rcTopTabs.left, rcTopTabs.bottom, _
                    nLeftWidth, AfxGetWindowHeight( HWND_FRMMAIN_TOPTABS_SHADOW ), _
                    SWP_NOZORDER or SWP_NOACTIVATE )

      ShowWindow( HWND_FRMMAIN_TOPTABS_SHADOW, nShow ) 

   else
      ShowWindow( HWND_FRMMAIN_TOPTABS_SHADOW, SW_HIDE )
   end if   

   function = 0
end function


' ========================================================================================
' Adjust all of the Tab rects left and right dimensions by the incoming amount
' ========================================================================================
function frmTopTabs_AdjustTabRects( byval nAdjAmount as long ) as long
   ' nAdjAmount can be +/-
   dim pDoc as clsDocument ptr
   
   for i as long = lbound(gTTabCtl.tabs) to ubound(gTTabCtl.tabs) 
      pDoc = gTTabCtl.tabs(i).pDoc
      if pDoc then
         gTTabCtl.tabs(i).rcTab.left += nAdjAmount
         gTTabCtl.tabs(i).rcTab.right += nAdjAmount
         gTTabCtl.tabs(i).rcIcon.left += nAdjAmount
         gTTabCtl.tabs(i).rcIcon.right += nAdjAmount
         gTTabCtl.tabs(i).rcText.left += nAdjAmount
         gTTabCtl.tabs(i).rcText.right += nAdjAmount
         gTTabCtl.tabs(i).rcClose.left += nAdjAmount
         gTTabCtl.tabs(i).rcClose.right += nAdjAmount
      end if
   next   
      
   function = 0
end function


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
function frmTopTabs_PositionWindows() As LRESULT

   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND_FRMMAIN_TOPTABS)
   If pWindow = 0 Then Exit Function
   
   ' Get the entire client area
   Dim As Rect rc: GetClientRect( HWND_FRMMAIN_TOPTABS, @rc )
   
   ' Calculate the RECT positions for all of the top tabs
   dim pDoc as clsDocument ptr 
   dim as long nTextWidth = 0
   dim as long nLeft = rc.left 
   dim as long TabHeight = TOPTABS_HEIGHT
   dim as long rcTextWidth = 70
   dim as long rcCloseWidth = 20
   dim as long rcIconWidth = 20   ' currently not using an icon in the tabs
   dim as long rightBorderWidth = 2
   dim as long hmargin = 3
   dim as long ActionPanelWidth = 114
   dim as long ActionButtonWidth = 24
   dim as long ActionButtonHeight = 22
   dim as long PrevTabsButtonWidth = 20
   dim as long PrevTabsButtonHeight = 20
   
   ' calculate the right hand side Action area panel that holds the
   ' more actions "..." button.
   gTTabCtl.rcActionPanel = rc
   gTTabCtl.rcActionPanel.left = gTTabCtl.rcActionPanel.right - pWindow->ScaleX( ActionPanelWidth )
   
   ' calculate the Prev and Next Tabs buttons
   dim as long vmargin = (TabHeight - PrevTabsButtonHeight) / 2
   gTTabCtl.rcPrevTabs.left = gTTabCtl.rcActionPanel.left + pWindow->ScaleX( hmargin * 2)
   gTTabCtl.rcPrevTabs.top = rc.top + pWindow->ScaleY(vmargin)
   gTTabCtl.rcPrevTabs.bottom = gTTabCtl.rcPrevTabs.top + pWindow->ScaleY( PrevTabsButtonHeight )
   gTTabCtl.rcPrevTabs.right = gTTabCtl.rcPrevTabs.left + pWindow->ScaleX( PrevTabsButtonWidth )

   gTTabCtl.rcNextTabs.left = gTTabCtl.rcPrevTabs.right + pWindow->ScaleX( hmargin )
   gTTabCtl.rcNextTabs.top = rc.top + pWindow->ScaleY(vmargin)
   gTTabCtl.rcNextTabs.bottom = gTTabCtl.rcNextTabs.top + pWindow->ScaleY( PrevTabsButtonHeight )
   gTTabCtl.rcNextTabs.right = gTTabCtl.rcNextTabs.left + pWindow->ScaleX( PrevTabsButtonWidth )

   ' calculate the split editor button
   gTTabCtl.rcSplitEditor.left = gTTabCtl.rcNextTabs.right + pWindow->ScaleX( hmargin )
   gTTabCtl.rcSplitEditor.top = rc.top + pWindow->ScaleY(vmargin)
   gTTabCtl.rcSplitEditor.bottom = gTTabCtl.rcSplitEditor.top + pWindow->ScaleY( PrevTabsButtonHeight )
   gTTabCtl.rcSplitEditor.right = gTTabCtl.rcSplitEditor.left + pWindow->ScaleX( PrevTabsButtonWidth )

   ' calculate the actual more actions "..." button itself
   vmargin = (TabHeight - ActionButtonHeight) / 2
   gTTabCtl.rcActionButton.left = gTTabCtl.rcSplitEditor.right + pWindow->ScaleX( hmargin )
   gTTabCtl.rcActionButton.top = rc.top + pWindow->ScaleY(vmargin)
   gTTabCtl.rcActionButton.bottom = gTTabCtl.rcActionButton.top + pWindow->ScaleY( ActionButtonHeight )
   gTTabCtl.rcActionButton.right = gTTabCtl.rcActionButton.left + pWindow->ScaleY( ActionButtonWidth )

   for i as long = lbound(gTTabCtl.tabs) to ubound(gTTabCtl.tabs) 
      pDoc = gTTabCtl.tabs(i).pDoc
      if pDoc then
         ' Determine the length of the text
         gTTabCtl.tabs(i).wszText = AfxStrPathname( "NAMEX", pDoc->DiskFilename )
         nTextWidth = getTextWidth( HWND_FRMMAIN_TOPTABS, gTTabCtl.tabs(i).wszText, ghMenuBar.hFontMenuBar, 0 )
         if nTextWidth < 70 then nTextWidth = 70

         ' calculate the various tab dimensions
         gTTabCtl.tabs(i).rcTab = rc
         gTTabCtl.tabs(i).rcTab.Left = nLeft
         gTTabCtl.tabs(i).rcTab.Right = nLeft + pWindow->ScaleX(rcIconWidth + nTextWidth + hmargin + rcCloseWidth + hmargin + rightBorderWidth)
         
         gTTabCtl.tabs(i).rcIcon = gTTabCtl.tabs(i).rcTab
         gTTabCtl.tabs(i).rcIcon.top = gTTabCtl.tabs(i).rcIcon.top + pWindow->ScaleY(vmargin)
         gTTabCtl.tabs(i).rcIcon.bottom = gTTabCtl.tabs(i).rcIcon.top + pWindow->ScaleY(rcCloseWidth)
         gTTabCtl.tabs(i).rcIcon.Left = gTTabCtl.tabs(i).rcTab.Left 
         gTTabCtl.tabs(i).rcIcon.Right = gTTabCtl.tabs(i).rcIcon.Left + pWindow->ScaleX(rcIconWidth)
         
         gTTabCtl.tabs(i).rcText = gTTabCtl.tabs(i).rcTab
         gTTabCtl.tabs(i).rcText.Left = gTTabCtl.tabs(i).rcIcon.Right
         gTTabCtl.tabs(i).rcText.Right = gTTabCtl.tabs(i).rcText.Left + pWindow->ScaleX(nTextWidth)
         
         gTTabCtl.tabs(i).rcClose = gTTabCtl.tabs(i).rcTab
         gTTabCtl.tabs(i).rcClose.top = gTTabCtl.tabs(i).rcClose.top + pWindow->ScaleY(vmargin)
         gTTabCtl.tabs(i).rcClose.bottom = gTTabCtl.tabs(i).rcClose.top + pWindow->ScaleY(rcCloseWidth)
         gTTabCtl.tabs(i).rcClose.Left = gTTabCtl.tabs(i).rcText.Right + hmargin + hmargin
         gTTabCtl.tabs(i).rcClose.Right = gTTabCtl.tabs(i).rcClose.Left + pWindow->ScaleX(rcCloseWidth)

         gTTabCtl.tabs(i).pDoc = pDoc
         gTTabCtl.tabs(i).isHot = false
         nLeft = gTTabCtl.tabs(i).rcTab.Right
      end if
   next

   ' We now know the exact width and postions of all tabs so now we need to calculate
   ' and adjust the gTTabCtl.FirstDisplayTab based on where the gTTabCtl.CurSel is located.
   if gTTabCtl.GetItemCount = 0 then
      ShowWindow( HWND_FRMMAIN_TOPTABS_SHADOW, SW_HIDE )
   else
      ' Adjust all of the Rects based on the current starting tab position
      dim as long nAdjAmount = gTTabCtl.tabs(gTTabCtl.FirstDisplayTab).rcTab.left
      frmTopTabs_AdjustTabRects( -(nAdjAmount) )
      
      gTTabCtl.ClientRightEdge = gTTabCtl.rcActionPanel.left
      
      if gTTabCtl.CurSel = -1 then
         ' Put If check for CurSel being invalid otherwise other tab calculations
         ' Scenarios will fail via GPF on invalid array access.
         
      ' SCENARIO #1
      ' New gTTabCtl.CurSel is located before gTTabCtl.FirstDisplayTab so it is visually
      ' not on screen. We simply need to move the gTTabCtl.CurSel rect into the
      ' the first position (gTTabCtl.FirstDisplayTab) and adjust all rects accordingly.
      elseif gTTabCtl.CurSel < gTTabCtl.FirstDisplayTab then
         ' The amount of the adjustment is equal to the difference between the
         ' current first tab's left edge and the new cursel's left edge.
         nAdjAmount = gTTabCtl.tabs(gTTabCtl.FirstDisplayTab).rcTab.left - gTTabCtl.tabs(gTTabCtl.CurSel).rcTab.left
         frmTopTabs_AdjustTabRects( nAdjAmount )
         gTTabCtl.FirstDisplayTab = gTTabCtl.CurSel
      
      ' SCENARIO #2
      ' The gTTabCtl.CurSel is already located within view but the right client edge
      ' for the tabcontrol may have grown wider so if the last tab's (ubound(gTTabCtl.tabs)
      ' right edge falls before the right edge then there may be an opportunity to shift 
      ' all of the tabs right to help fill the space.
      elseif gTTabCtl.tabs(gTTabCtl.CurSel).rcTab.right < gTTabCtl.ClientRightEdge then
         if gTTabCtl.tabs(ubound(gTTabCtl.tabs)).rcTab.right < gTTabCtl.ClientRightEdge then
            nAdjAmount = 0
            if gTTabCtl.FirstDisplayTab - 1 >= 0 then
               dim as long nTabWidth = _
                       gTTabCtl.tabs(gTTabCtl.FirstDisplayTab - 1).rcTab.Right - gTTabCtl.tabs(gTTabCtl.FirstDisplayTab - 1).rcTab.left
               if gTTabCtl.tabs(ubound(gTTabCtl.tabs)).rcTab.right + nTabWidth <= gTTabCtl.ClientRightEdge then
                  frmTopTabs_AdjustTabRects( nTabWidth )
                  gTTabCtl.FirstDisplayTab = gTTabCtl.FirstDisplayTab - 1
               end if   
            end if
         end if
      
      ' SCENARIO #3
      ' gTTabCtl.CurSel is located past the tab control's right edge. We can begin
      ' "removing" tabs at the start of the view until the gTTabCtl.CurSel tab comes
      ' completely within view.
      elseif gTTabCtl.tabs(gTTabCtl.CurSel).rcTab.right > gTTabCtl.ClientRightEdge then
         ' need to determine how far to move the gTTabCtl.FirstDisplayTab
         nAdjAmount = 0
         for i as long = gTTabCtl.FirstDisplayTab to gTTabCtl.CurSel 
            ' remove the width of each successive tab starting with the
            ' first tab and work our way until we reach the target tab.
            dim as long nTabWidth = gTTabCtl.tabs(i).rcTab.Right - gTTabCtl.tabs(i).rcTab.left
            nAdjAmount += nTabWidth
            if gTTabCtl.tabs(gTTabCtl.CurSel).rcTab.Right - nAdjAmount <= gTTabCtl.ClientRightEdge then
               gTTabCtl.FirstDisplayTab = i + 1
               frmTopTabs_AdjustTabRects( -(nAdjAmount) )
               exit for
            end if   
         next
      end if

      ' sanity checks
      if gTTabCtl.FirstDisplayTab > ubound(gTTabCtl.tabs) then 
         gTTabCtl.FirstDisplayTab = ubound(gTTabCtl.tabs)
      elseif gTTabCtl.FirstDisplayTab < 0 then 
         gTTabCtl.FirstDisplayTab = 0 
      end if   
   
   end if
   AfxRedrawWindow( HWND_FRMMAIN_TOPTABS )

   Function = 0
End Function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmTopTabs
' ========================================================================================
Function frmTopTabs_OnSize( _
            ByVal HWnd As HWnd, _
            ByVal state As UINT, _
            ByVal cx As Long, _
            ByVal cy As Long _
            ) As LRESULT

   If state <> SIZE_MINIMIZED Then
      frmTopTabs_PositionWindows()
   End If

   Function = 0
End Function


' ========================================================================================
' Do hit test to determine what tab is currently under the mouse cursor
' ========================================================================================
function getHotTabHitTest( byval hWin as HWnd ) as long
   dim as POINT pt: GetCursorPos( @pt )
   MapWindowPoints( HWND_DESKTOP, hWin, cast( POINT ptr, @pt ), 1 )
   dim as long hotTab = -1
   for i as long = lbound(gTTabCtl.tabs) to ubound(gTTabCtl.tabs)
      if PtInRect( @gTTabCtl.tabs(i).rcTab, pt ) then
         hotTab = i
         gTTabCtl.tabs(i).isHot = true
      else   
         gTTabCtl.tabs(i).isHot = false
      end if
   next
   function = hotTab
end function

' ========================================================================================
' Do hit test to determine if "X" close button on tab was clicked
' ========================================================================================
function isTabCloseHitTest( byval hWin as HWnd, byval idx as long ) as long
   if gTTabCtl.IsSafeIndex(idx) = false then exit function
   dim as POINT pt: GetCursorPos( @pt )
   MapWindowPoints( HWND_DESKTOP, hWin, cast( POINT ptr, @pt ), 1 )
   if PtInRect( @gTTabCtl.tabs(idx).rcClose, pt ) then return idx
   function = -1
end function

' ========================================================================================
' Do hit test to determine if "..." button in action Area was clicked
' ========================================================================================
function isActionButtonHitTest( byval hWin as HWnd ) as boolean
   dim as POINT pt: GetCursorPos( @pt )
   MapWindowPoints( HWND_DESKTOP, hWin, cast( POINT ptr, @pt ), 1 )
   if PtInRect( @gTTabCtl.rcActionButton, pt ) then return true
   function = false
end function

' ========================================================================================
' Do hit test to determine if Prev Tabs button in action Area was clicked
' ========================================================================================
function isPrevTabsHitTest( byval hWin as HWnd ) as boolean
   dim as POINT pt: GetCursorPos( @pt )
   MapWindowPoints( HWND_DESKTOP, hWin, cast( POINT ptr, @pt ), 1 )
   if PtInRect( @gTTabCtl.rcPrevTabs, pt ) then return true
   function = false
end function

' ========================================================================================
' Do hit test to determine if Next Tabs button in action Area was clicked
' ========================================================================================
function isNextTabsHitTest( byval hWin as HWnd ) as boolean
   dim as POINT pt: GetCursorPos( @pt )
   MapWindowPoints( HWND_DESKTOP, hWin, cast( POINT ptr, @pt ), 1 )
   if PtInRect( @gTTabCtl.rcNextTabs, pt ) then return true
   function = false
end function

' ========================================================================================
' Do hit test to determine if SplitEditor button in action Area was clicked
' ========================================================================================
function isSplitEditorHitTest( byval hWin as HWnd ) as boolean
   dim as POINT pt: GetCursorPos( @pt )
   MapWindowPoints( HWND_DESKTOP, hWin, cast( POINT ptr, @pt ), 1 )
   if PtInRect( @gTTabCtl.rcSplitEditor, pt ) then return true
   function = false
end function

' ========================================================================================
' Do hit test if over the Action Panel (no action but prevents other tooltips)
' ========================================================================================
function isActionPanelHitTest( byval hWin as HWnd ) as boolean
   dim as POINT pt: GetCursorPos( @pt )
   MapWindowPoints( HWND_DESKTOP, hWin, cast( POINT ptr, @pt ), 1 )
   if PtInRect( @gTTabCtl.rcActionPanel, pt ) then return true
   function = false
end function

' ========================================================================================
' frmTopTabsShadow_WndProc Window procedure
' ========================================================================================
Function frmTopTabsShadow_WndProc( _
            ByVal HWnd   As HWnd, _
            ByVal uMsg   As UINT, _
            ByVal wParam As WPARAM, _
            ByVal lParam As LPARAM _
            ) As LRESULT

   Select Case uMsg
      case WM_DESTROY
         DIM pWindow AS CWindow PTR = AfxCWindowPtr(HWnd)
         If pWindow = 0 Then Delete pWindow
   end select      

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

end function

' ========================================================================================
' frmTopTabs Window procedure
' ========================================================================================
Function frmTopTabs_WndProc( _
            ByVal HWnd   As HWnd, _
            ByVal uMsg   As UINT, _
            ByVal wParam As WPARAM, _
            ByVal lParam As LPARAM _
            ) As LRESULT

   static as long curr_idxHot = -1
   static hTooltip as HWND
   
   Select Case uMsg
      HANDLE_MSG (HWnd, WM_SIZE, frmTopTabs_OnSize)
   
   case WM_DESTROY
      DIM pWindow AS CWindow PTR = AfxCWindowPtr(HWnd)
      If pWindow = 0 Then Delete pWindow
      DestroyWindow( HWND_FRMMAIN_TOPTABS_SHADOW )
      
   case WM_ERASEBKGND
      return true

   case WM_MOUSEMOVE
      Dim tme As TrackMouseEvent
      tme.cbSize = Sizeof(TrackMouseEvent)
      tme.dwFlags = TME_HOVER Or TME_LEAVE
      tme.hwndTrack = HWnd
      TrackMouseEvent(@tme) 

      if IsWindow(hTooltip) = 0 then hTooltip = AfxAddTooltip( HWnd, "", false, false )

      curr_idxHot = getHotTabHitTest( HWnd )
      AfxRedrawWindow( HWnd )
      
      if gApp.bDragTabActive then
         if gTTabCtl.IsSafeIndex(curr_idxHot) = false then exit function
         SetCursor( LoadCursor(0, MAKEINTRESOURCE(OCR_SIZEALL)) )
         ' Swap the two elements in the array
         swap gTTabCtl.tabs(gTTabCtl.CurSel), gTTabCtl.tabs(curr_idxHot)
         gTTabCtl.SetFocusTab( curr_idxHot )
         frmTopTabs_PositionWindows()
      end if
      
      
   case WM_MOUSELEAVE
      ' this will reset all tabs isHot to -1 and curr_idxHot to -1
      curr_idxHot = getHotTabHitTest( HWnd )
      AfxDeleteTooltip( hTooltip, HWnd )
      hTooltip = 0
      AfxRedrawWindow( HWnd )


   case WM_MOUSEHOVER
      dim as CWSTR wszTooltip 
      ' test for Action Button, Filename caption or Close caption
      curr_idxHot = getHotTabHitTest( HWnd )
      if isPrevTabsHitTest( HWnd ) then
         wszTooltip  = L(431, "Scroll Tabs Left")
      elseif isNextTabsHitTest( HWnd ) then
         wszTooltip  = L(432, "Scroll Tabs Right")
      elseif isSplitEditorHitTest( HWnd ) then
         wszTooltip  = L(447, "Split Editor")
      elseif isActionButtonHitTest( HWnd ) then
         wszTooltip  = L(430, "Tab List")
      elseif isActionPanelHitTest( HWnd ) then
         ' do nothing. no tooltip.
      elseif isTabCloseHitTest( HWnd, curr_idxHot ) = curr_idxHot then
         wszTooltip  = L(86, "Close Tab")
      else
         if gTTabCtl.IsSafeIndex(curr_idxHot) = false then exit function
         wszTooltip = gTTabCtl.tabs(curr_idxHot).pDoc->DiskFilename   
      end if
      ' Display the tooltip
      AfxSetTooltipText( hTooltip, HWnd, wszTooltip )
      AfxRedrawWindow( HWnd )


   case WM_RBUTTONDOWN
      if isPrevTabsHitTest( HWnd ) then
      elseif isNextTabsHitTest( HWnd ) then
      elseif isSplitEditorHitTest( HWnd ) then
      elseif isActionButtonHitTest( HWnd ) then
      elseif isActionPanelHitTest( HWnd ) then
      else
         ' Create the popup menu
         curr_idxHot = getHotTabHitTest( HWnd )
         if gTTabCtl.IsSafeIndex(curr_idxHot) = false then exit function
         gTTabCtl.SetFocusTab(curr_idxHot)
         Dim As POINT pt
         dim as HMENU hPopupMenu = CreateTopTabCtlContextMenu(curr_idxHot)
         GetCursorPos @pt
         TrackPopupMenu(hPopUpMenu, 0, pt.x, pt.y, 0, HWND_FRMMAIN, ByVal Null)
         DestroyMenu hPopUpMenu
         Return True   ' prevent further processing that leads to WM_CONTEXTMENU
      end if

   case WM_LBUTTONDOWN
      if isPrevTabsHitTest( HWnd ) then
      elseif isNextTabsHitTest( HWnd ) then
      elseif isSplitEditorHitTest( HWnd ) then
      elseif isActionButtonHitTest( HWnd ) then
      elseif isActionPanelHitTest( HWnd ) then
      else
         curr_idxHot = getHotTabHitTest( HWnd )
         if gTTabCtl.IsSafeIndex(curr_idxHot) = false then exit function
         if (isTabCloseHitTest( HWnd, curr_idxHot ) = curr_idxHot) and (curr_idxHot <> gTTabCtl.CurSel) then
            ' we are mouse down on an "X" close on a tab that is not currently the active tab. We
            ' will simply close that tab rather than bring it into focus and then delete.
            ' allow the mouse message to eventually bubble up WM_LBUTTONUP where we do the close.
         else
            ' we have clicked on tab so bring it into focus
            gTTabCtl.SetFocusTab( curr_idxHot )
            if (wParam and MK_SHIFT) then
               gApp.bDragTabActive = true
               SetCursor( LoadCursor(0, MAKEINTRESOURCE(OCR_SIZEALL)) )
            end if
            AfxRedrawWindow( HWnd )
         end if
         SetCapture(hWnd)
      end if

   case WM_LBUTTONDBLCLK
      ' This will have already fired WM_LBUTTONUP
      ' If the tab holds a pDoc that is a visual designer form then we will toggle between
      ' code and design by toggling the designer tab control.
      dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
      if pDoc then
         if pDoc->IsDesigner then
            pDoc->DesignTabsCurSel = iif(pDoc->DesignTabsCurSel = 0, 1, 0)
            frmMain_PositionWindows
            PostMessage( HWND_FRMMAIN, MSG_USER_GENERATECODE, 0, 0 )
         end if
      end if
      
   case WM_LBUTTONUP
      if isPrevTabsHitTest( HWnd ) then
         ' if the first displayed tab is not index 0 then there must be tabs before it
         ' that can be displayed. If no previous tabs then set focus to the first tab.
         if gTTabCtl.FirstDisplayTab > 0 then
            gTTabCtl.SetFocusTab( gTTabCtl.FirstDisplayTab - 1 )
         else   
            gTTabCtl.SetFocusTab(0)
         end if
      elseif isNextTabsHitTest( HWnd ) then
         ' determine if there are any tabs to the right of the current display view. Bring
         ' the first one found (if any) into view. Keep track of the found tab because if
         ' none found then simply set the CurSel to the last tab.
         dim as long nFound = -1
         for i as long = gTTabCtl.FirstDisplayTab to gTTabCtl.GetItemCount - 1
            if gTTabCtl.tabs(i).rcTab.Left > gTTabCtl.ClientRightEdge then
               nFound = i: exit for
               exit for
            end if
         next
         if nFound = -1 then nFound = gTTabCtl.GetItemCount - 1
         gTTabCtl.SetFocusTab(nFound)

      elseif isSplitEditorHitTest( HWnd ) then
         dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
         if pDoc then
            pDoc->bEditorIsSplit = not pDoc->bEditorIsSplit
            if pDoc->bEditorIsSplit then
               dim as RECT rc = AfxGetWindowRect(pDoc->hWindow(0))
               MapWindowPoints( HWND_DESKTOP, HWND_FRMMAIN, cast(POINT ptr, @rc), 2 )
               pDoc->SplitY = (rc.bottom - rc.top) / 2
            else
               pDoc->SplitY = 0
            end if      
            frmMain_PositionWindows()
         end if
         
      elseif isActionButtonHitTest( HWnd ) then
            dim as HMENU hPopupMenu = CreateTopTabsActionButtonContextMenu()
            ' Popup the menu to the bottom of the Action Button (right aligned)
            dim as RECT rc = gTTabCtl.rcActionButton   ' work with a copy
            MapWindowPoints( HWND_FRMMAIN_TOPTABS, HWND_DESKTOP, cast(POINT ptr, @rc), 2 )
            dim as long id = TrackPopupMenu(hPopUpMenu, _
                                TPM_RIGHTALIGN or TPM_RETURNCMD, _
                                rc.right, rc.bottom, 0, HWND_FRMMAIN, byval null)
            ' Return value is 1 based because 0 indicates that menu was cancelled
            if (id > 0) andalso (id <= gTTabCtl.GetItemCount) then
               gTTabCtl.SetFocusTab(id-1)
            end if
            DestroyMenu( hPopUpMenu )
            Return true   ' prevent further processing that leads to WM_CONTEXTMENU

      elseif isActionPanelHitTest( HWnd ) then
      else
         ' Reset the mouse pointer
         SetCursor( LoadCursor( null, IDC_ARROW ))
         ReleaseCapture
         gApp.bDragTabActive = false
         ' If we are still over an "X" close during buttonup then close the tab
         curr_idxHot = getHotTabHitTest( HWnd )
         if gTTabCtl.IsSafeIndex(curr_idxHot) = false then exit function
         if isTabCloseHitTest( HWnd, curr_idxHot ) = curr_idxHot then
            gTTabCtl.CloseTab( curr_idxHot )
            frmTopTabs_PositionWindows()
         end if
      end if

   case WM_PAINT
      Dim As PAINTSTRUCT ps
      Dim As HDC hDc

      Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND_FRMMAIN_TOPTABS)
      If pWindow = 0 Then Exit Function

      hDC = BeginPaint( hWnd, @ps )

      SaveDC(hDC)
      dim as long nWidth = ps.rcPaint.right - ps.rcPaint.left
      dim as long nHeight = ps.rcPaint.bottom - ps.rcPaint.top

      Dim memDC as HDC      ' Double buffering
      Dim hbit As HBITMAP   ' Double buffering

      memDC = CreateCompatibleDC( hDC )
      hbit  = CreateCompatibleBitmap( hDC, nWidth, nHeight )
      If hbit Then hbit = SelectObject( memDC, hbit )

      ' Fill in the entire back panel width across the top of the screen
      FillRect( memDC, @ps.rcPaint, ghTopTabs.hPanelBrush )

      ' Create a black pen that acts as the divider for each tab
      dim as long penWidth = pWindow->ScaleX(1)
      dim as HPEN hPenSolid = CreatePen( PS_SOLID, penWidth, ghTopTabs.Divider )
      SelectObject( memDC, hPenSolid)
      
      dim as HPEN hPenNull = CreatePen( PS_NULL, 1, 0 )  ' null/invisible pen

      dim as POINT pt: GetCursorPos( @pt )
      dim as RECT rc
      
      ' All of the rc calculations have already been done in frmTopTabs_PostionWindows
      if gTTabCtl.FirstDisplayTab > ubound(gTTabCtl.tabs) then gTTabCtl.FirstDisplayTab = lbound(gTTabCtl.tabs)
      for i as long = gTTabCtl.FirstDisplayTab to ubound(gTTabCtl.tabs)
         ' paint this tab based on active/inactive status
         if i = gTTabCtl.CurSel then
            SetBkColor( memDC, ghTopTabs.BackColorHot )
            SetTextColor( memDC, ghTopTabs.ForeColorHot )
            FillRect( memDC, @gTTabCtl.tabs(i).rcTab, ghTopTabs.hBackBrushHot )
         else
            SetBkColor( memDC, ghTopTabs.BackColor )
            SetTextColor( memDC, ghTopTabs.ForeColor )
            FillRect( memDC, @gTTabCtl.tabs(i).rcTab, ghTopTabs.hBackBrush )
         end if
         
         ' display the rcIcon
         'dim as long wsStyle = DT_NOPREFIX or DT_CENTER Or DT_VCENTER or DT_SINGLELINE
         'SelectObject( memDC, ghMenuBar.hFontSymbolSmall )
         'DrawText( memDC, wszTabDocument, -1, Cast(lpRect, @gTTabCtl.tabs(i).rcIcon), wsStyle )
         
         ' display the tab text
         dim as long wsStyle = DT_NOPREFIX or DT_CENTER Or DT_VCENTER or DT_SINGLELINE
         SelectObject(memDC, ghMenuBar.hFontMenuBar)
         DrawText( memDC, gTTabCtl.tabs(i).wszText.sptr, -1, Cast(lpRect, @gTTabCtl.tabs(i).rcText), wsStyle )
         
         ' if this document is dirty then the dirty circle icon overrides the display of the close icon
         ' unless the tab isHot. The dirty icon displays always for active and inactive tabs.
         dim as boolean isDirty 
         If cbool(SciExec( gTTabCtl.tabs(i).pDoc->hWindow(0), SCI_GETMODIFY, 0, 0 )) then isDirty = true
         if gTTabCtl.tabs(i).pDoc->UserModified then isDirty = true
         
         if isDirty then
            ' if mouse is over the rcClose then draw the close icon, otherwise draw the dirty circle
            SelectObject( memDC, ghMenuBar.hFontSymbolSmall )
            DrawText( memDC, wszCompileResultIcon, -1, Cast(lpRect, @gTTabCtl.tabs(i).rcClose), wsStyle )
         end if
         
         if (i = gTTabCtl.CurSel) or (gTTabCtl.tabs(i).isHot = true) then
            ' if this is the active tab or it is mouse hot then display the close icon only 
            ' if the dirty icon is not already being displayed
            rc = gTTabCtl.tabs(i).rcClose
            if isDirty = false then
               SelectObject( memDC, ghMenuBar.hFontSymbolSmall )
               DrawText( memDC, wszClose, -1, Cast(lpRect, @rc), wsStyle )
            end if
            if (isTabCloseHitTest( HWnd, i ) = i) andalso (gApp.bDragTabActive = false) then
               ' if we are hovered over the "X" close icon rect then highlight it
               SelectPen( memDC, hPenNull )
               SelectObject( memDC, ghTopTabs.hCloseBrushHot )
               RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
               SetBkColor( memDC, ghTopTabs.CloseBackColorHot )
               SelectObject( memDC, ghMenuBar.hFontSymbolSmall )
               DrawText( memDC, wszClose, -1, Cast(lpRect, @rc), wsStyle )
            end if   
         end if
         
         ' Draw the righthand side black divider
         SelectPen( memDC, hPenSolid )
         MoveToEx( memDC, gTTabCtl.tabs(i).rcTab.Right - penWidth, gTTabCtl.tabs(i).rcTab.top, null )
         LineTo( memDC, gTTabCtl.tabs(i).rcTab.Right - penWidth, gTTabCtl.tabs(i).rcTab.bottom )   
         
      next
      
      ' Lastly, draw the TopTabs Action Area (basically, the little panel to the
      ' right of all of the top tabs that has the "..." icon.
      if gTTabCtl.GetItemCount then
         ' paint the full Action Panel
         FillRect( memDC, @gTTabCtl.rcActionPanel, ghTopTabs.hPanelBrush )
         
         ' Draw the PrevTabs button
         SelectObject( memDC, ghTopTabs.hPanelBrush )
         SetBkColor( memDC, ghTopTabs.BackColor )
         SetTextColor( memDC, ghTopTabs.ForeColorHot )
         if isPrevTabsHitTest( HWnd ) then
            SetBkColor( memDC, ghTopTabs.CloseBackColorHot )
            SelectObject( memDC, ghTopTabs.hCloseBrushHot )
         end if
         rc = gTTabCtl.rcPrevTabs
         SelectPen( memDC, hPenNull )
         RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
         dim as long wsStyle = DT_NOPREFIX or DT_CENTER Or DT_TOP 'or DT_SINGLELINE
         SelectObject(memDC, ghMenuBar.hFontSymbol)
         DrawText( memDC, wszChevronLeft, -1, Cast(lpRect, @rc), wsStyle )

         ' Draw the NextTabs button
         SelectObject( memDC, ghTopTabs.hPanelBrush )
         SetBkColor( memDC, ghTopTabs.BackColor )
         SetTextColor( memDC, ghTopTabs.ForeColorHot )
         if isNextTabsHitTest( HWnd ) then
            SetBkColor( memDC, ghTopTabs.CloseBackColorHot )
            SelectObject( memDC, ghTopTabs.hCloseBrushHot )
         end if
         rc = gTTabCtl.rcNextTabs
         SelectPen( memDC, hPenNull )
         RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
         wsStyle = DT_NOPREFIX or DT_CENTER Or DT_TOP 'or DT_SINGLELINE
         SelectObject(memDC, ghMenuBar.hFontSymbol)
         DrawText( memDC, wszChevronRight, -1, Cast(lpRect, @rc), wsStyle )
         
         ' Draw the SplitEditor button
         SelectObject( memDC, ghTopTabs.hPanelBrush )
         SetBkColor( memDC, ghTopTabs.BackColor )
         SetTextColor( memDC, ghTopTabs.ForeColorHot )
         if isSplitEditorHitTest( HWnd ) then
            SetBkColor( memDC, ghTopTabs.CloseBackColorHot )
            SelectObject( memDC, ghTopTabs.hCloseBrushHot )
         end if
         rc = gTTabCtl.rcSplitEditor
         SelectPen( memDC, hPenNull )
         RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
         wsStyle = DT_NOPREFIX or DT_CENTER Or DT_TOP 'or DT_SINGLELINE
         SelectObject(memDC, ghMenuBar.hFontSymbol)
         DrawText( memDC, wszSplitEditor, -1, Cast(lpRect, @rc), wsStyle )

         ' Draw the Action buttom "..."
         SelectObject( memDC, ghTopTabs.hPanelBrush )
         SetBkColor( memDC, ghTopTabs.BackColor )
         SetTextColor( memDC, ghTopTabs.ForeColorHot )
         if isActionButtonHitTest( HWnd ) then
            SetBkColor( memDC, ghTopTabs.CloseBackColorHot )
            SelectObject( memDC, ghTopTabs.hCloseBrushHot )
         end if
         rc = gTTabCtl.rcActionButton
         SelectPen( memDC, hPenNull )
         RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
         wsStyle = DT_NOPREFIX or DT_CENTER Or DT_VCENTER 'or DT_SINGLELINE
         SelectObject(memDC, ghMenuBar.hFontSymbolLargeBold)
         DrawText( memDC, wszMoreActions, -1, Cast(lpRect, @rc), wsStyle )
      end if
   
      BitBlt( hDC, 0, 0, nWidth, nHeight, memDC, 0, 0, SRCCOPY )

      ' Cleanup
      if hPenSolid then DeleteObject( hPenSolid )
      if hPenNull  then DeleteObject( hPenNull )
      if hbit  then DeleteObject( SelectObject(memDC, hbit) )
      if memDC then DeleteDC( memDC )
      RestoreDC( hDC, -1 )

      EndPaint( hWnd, @ps )
   
   End Select
   
   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function



' ========================================================================================
' frmTopTabs_Show
' ========================================================================================
Function frmTopTabs_Show( ByVal hWndParent As HWnd ) As LRESULT

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   HWND_FRMMAIN_TOPTABS = pWindow->Create( hWndParent, "", @frmTopTabs_WndProc, _
        0, 0, 0, TOPTABS_HEIGHT, _
        WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   ' Disable background erasing by only assigning the one style
   pWindow->ClassStyle = CS_DBLCLKS

   ' create semi-transparent window offset under our tabcontrol in order to simulate a shadow.
   if gApp.isWineActive = false then
      pWindow = New CWindow
      pWindow->DPI = AfxCWindowPtr(HWND_FRMMAIN)->DPI
      HWND_FRMMAIN_TOPTABS_SHADOW = pWindow->Create( HWND_FRMMAIN, "", _
           @frmTopTabsShadow_WndProc, 0, 0, 0, pWindow->ScaleY(1), _
           WS_POPUP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, WS_EX_LAYERED or WS_EX_NOACTIVATE )
      pWindow->ClassStyle = CS_DBLCLKS
      pWindow->Brush = GetSysColorBrush(COLOR_WINDOWTEXT + 1)   ' black background
      ' 0 totally transparent, 255 totally opaque
      SetLayeredWindowAttributes( HWND_FRMMAIN_TOPTABS_SHADOW, GetSysColor(COLOR_WINDOWTEXT + 1) , 100, LWA_ALPHA ) 
   end if
   
   frmTopTabs_PositionWindows()
   Function = 0
   
End Function

